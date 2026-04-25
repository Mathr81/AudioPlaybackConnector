#pragma warning(disable:4819)
#include "pch.h"
#include "AudioPlaybackConnector.h"
#include "AudioUtil.hpp"

using namespace winrt::Windows::Data::Json;
using namespace winrt::Windows::Devices::Enumeration;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Media::Audio;
using namespace winrt::Windows::UI::Xaml;
using namespace winrt::Windows::UI::Xaml::Controls;
using namespace winrt::Windows::UI::Xaml::Hosting;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
// ... (rest of function declarations)
void SetupFlyout();
void SetupMenu();
winrt::fire_and_forget ConnectDevice(DevicePicker, std::wstring);
void SetupDevicePicker();
void SetupSvgIcon();
void UpdateNotifyIcon();
bool GetStartupStatus();
void SetStartupStatus(bool status);
void ShowInitialToastNotification();
bool TryGetWorkerDeviceId(std::wstring& deviceId);
bool TryGetArgValue(PCWSTR name, std::wstring& value);
int RunWorkerProcess(std::wstring_view deviceId, std::wstring_view stopEventName, std::wstring_view workerAppId);
bool LaunchWorkerProcess(std::wstring_view deviceId, WorkerProcessInfo& workerInfo);
void PruneExitedWorkers();
void RequestStopWorker(WorkerProcessInfo& workerInfo, DWORD waitMs);
void StopAndCleanupWorker(WorkerProcessInfo& workerInfo);
fs::path GetWorkerExecutablePath(std::wstring_view deviceId, uint64_t launchId);
bool EnsureWorkerExecutable(const fs::path& workerExePath);
std::wstring GetWorkerAppId(std::wstring_view deviceId, uint64_t launchId);

bool IsSystemLightTheme()
{
	wil::unique_hkey hKey;
	if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
	{
		DWORD value = 1;
		DWORD size = sizeof(value);
		DWORD type = REG_DWORD;
		if (RegQueryValueExW(hKey.get(), L"SystemUsesLightTheme", nullptr, &type, reinterpret_cast<LPBYTE>(&value), &size) == ERROR_SUCCESS && type == REG_DWORD)
		{
			return value != 0;
		}
	}
	return false;
}

HICON CreateNotifyIcon(size_t connectionCount)
{
	if (g_notifyIconSvg.empty())
	{
		return nullptr;
	}

	const int width = GetSystemMetrics(SM_CXSMICON);
	const int height = GetSystemMetrics(SM_CYSMICON);
	const bool hasConnection = connectionCount > 0;
	D2D1_COLOR_F baseColor = hasConnection ? D2D1::ColorF(0.0f, 0.47f, 1.0f, 1.0f) : (IsSystemLightTheme() ? D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f) : D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f));

	wil::unique_hdc hdc(CreateCompatibleDC(nullptr));
	THROW_IF_NULL_ALLOC(hdc);

	auto hBitmap = CreateDIB(hdc.get(), width, height, 32);
	THROW_IF_NULL_ALLOC(hBitmap);
	auto hBitmapMask = CreateDIB(hdc.get(), width, height, 1);
	THROW_IF_NULL_ALLOC(hBitmapMask);

	auto select = wil::SelectObject(hdc.get(), hBitmap.get());
	DrawSvgTohDC(g_notifyIconSvg, hdc.get(), width, height, baseColor);

	if (hasConnection)
	{
		const int badgeRadius = max(4, width / 4);
		const int margin = 1;
		const int cx = width - badgeRadius - margin;
		const int cy = height - badgeRadius - margin;
		RECT badgeRect = { cx - badgeRadius, cy - badgeRadius, cx + badgeRadius, cy + badgeRadius };

		wil::unique_hbrush badgeBrush(CreateSolidBrush(RGB(0, 120, 215)));
		auto oldBrush = SelectObject(hdc.get(), badgeBrush.get());
		auto oldPen = SelectObject(hdc.get(), GetStockObject(NULL_PEN));
		Ellipse(hdc.get(), badgeRect.left, badgeRect.top, badgeRect.right, badgeRect.bottom);
		SelectObject(hdc.get(), oldPen);
		SelectObject(hdc.get(), oldBrush);

		std::wstring text = connectionCount > 99 ? L"99+" : std::to_wstring(connectionCount);
		HFONT font = CreateFontW(-MulDiv(7, GetDeviceCaps(hdc.get(), LOGPIXELSY), 72), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, L"Segoe UI");
		if (font)
		{
			auto oldFont = SelectObject(hdc.get(), font);
			SetTextColor(hdc.get(), RGB(255, 255, 255));
			SetBkMode(hdc.get(), TRANSPARENT);
			DrawTextW(hdc.get(), text.c_str(), -1, &badgeRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
			SelectObject(hdc.get(), oldFont);
			DeleteObject(font);
		}

		DIBSECTION dib = {};
		if (GetObjectW(hBitmap.get(), sizeof(dib), &dib) == sizeof(dib) && dib.dsBm.bmBits)
		{
			auto pixels = static_cast<uint32_t*>(dib.dsBm.bmBits);
			const int pixelCount = dib.dsBm.bmWidth * abs(dib.dsBm.bmHeight);
			for (int i = 0; i < pixelCount; ++i)
			{
				if ((pixels[i] & 0xFF000000u) == 0 && (pixels[i] & 0x00FFFFFFu) != 0)
				{
					pixels[i] |= 0xFF000000u;
				}
			}
		}
	}

	ICONINFO iconInfo = {
		.fIcon = TRUE,
		.hbmMask = hBitmapMask.get(),
		.hbmColor = hBitmap.get()
	};

	HICON hIcon = CreateIconIndirect(&iconInfo);
	THROW_LAST_ERROR_IF_NULL(hIcon);
	return hIcon;
}

bool TryGetArgValue(PCWSTR name, std::wstring& value)
{
	int argc = 0;
	auto argv = CommandLineToArgvW(GetCommandLineW(), &argc);
	if (!argv)
	{
		return false;
	}

	for (int i = 1; i < argc; ++i)
	{
		if (_wcsicmp(argv[i], name) == 0 && i + 1 < argc)
		{
			value = argv[i + 1];
			LocalFree(argv);
			return true;
		}
	}

	LocalFree(argv);
	return false;
}

bool TryGetWorkerDeviceId(std::wstring& deviceId)
{
	return TryGetArgValue(L"--worker", deviceId);
}

void RequestStopWorker(WorkerProcessInfo& workerInfo, DWORD waitMs)
{
	if (workerInfo.stopEventHandle)
	{
		SetEvent(workerInfo.stopEventHandle);
	}
	if (workerInfo.processHandle)
	{
		if (WaitForSingleObject(workerInfo.processHandle, waitMs) != WAIT_OBJECT_0)
		{
			TerminateProcess(workerInfo.processHandle, 0);
			WaitForSingleObject(workerInfo.processHandle, 200);
		}
	}
}

void StopAndCleanupWorker(WorkerProcessInfo& workerInfo)
{
	RequestStopWorker(workerInfo, 2000);
	if (workerInfo.processHandle)
	{
		CloseHandle(workerInfo.processHandle);
		workerInfo.processHandle = nullptr;
	}
	if (workerInfo.stopEventHandle)
	{
		CloseHandle(workerInfo.stopEventHandle);
		workerInfo.stopEventHandle = nullptr;
	}
	if (!workerInfo.executablePath.empty())
	{
		DeleteFileW(workerInfo.executablePath.c_str());
		workerInfo.executablePath.clear();
	}
}

int RunWorkerProcess(std::wstring_view deviceId, std::wstring_view stopEventName, std::wstring_view workerAppId)
{
	try
	{
		if (!workerAppId.empty())
		{
			SetCurrentProcessExplicitAppUserModelID(std::wstring(workerAppId).c_str());
		}

		winrt::init_apartment();

		auto connection = AudioPlaybackConnection::TryCreateFromId(deviceId);
		if (!connection)
		{
			return EXIT_FAILURE;
		}

		auto hClosed = CreateEventW(nullptr, TRUE, FALSE, nullptr);
		if (!hClosed)
		{
			return EXIT_FAILURE;
		}

		HANDLE hStop = OpenEventW(SYNCHRONIZE, FALSE, std::wstring(stopEventName).c_str());
		if (!hStop)
		{
			CloseHandle(hClosed);
			return EXIT_FAILURE;
		}

		connection.StateChanged([hClosed](const auto& sender, const auto&) {
			if (sender.State() == AudioPlaybackConnectionState::Closed)
			{
				SetEvent(hClosed);
			}
		});

		connection.StartAsync().get();
		auto result = connection.OpenAsync().get();
		if (result.Status() != AudioPlaybackConnectionOpenResultStatus::Success)
		{
			CloseHandle(hStop);
			CloseHandle(hClosed);
			return EXIT_FAILURE;
		}

		HANDLE handles[2] = { hClosed, hStop };
		auto waitResult = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
		if (waitResult == WAIT_OBJECT_0 + 1)
		{
			// Stop requested: exit worker and let process teardown release this connection only.
		}

		CloseHandle(hStop);
		CloseHandle(hClosed);
	}
	catch (winrt::hresult_error const&)
	{
		LOG_CAUGHT_EXCEPTION();
		return EXIT_FAILURE;
	}

	return EXIT_SUCCESS;
}

bool LaunchWorkerProcess(std::wstring_view deviceId, WorkerProcessInfo& workerInfo)
{
	const uint64_t launchId = ++g_workerEventSerial;
	auto workerExePath = GetWorkerExecutablePath(deviceId, launchId);
	if (!EnsureWorkerExecutable(workerExePath))
	{
		return false;
	}

	std::wstring stopEventName = L"Local\\AudioPlaybackConnector_WorkerStop_" + std::to_wstring(GetCurrentProcessId()) + L"_" + std::to_wstring(launchId);
	std::wstring workerAppId = GetWorkerAppId(deviceId, launchId);

	workerInfo.stopEventHandle = CreateEventW(nullptr, TRUE, FALSE, stopEventName.c_str());
	if (!workerInfo.stopEventHandle)
	{
		LOG_LAST_ERROR();
		return false;
	}

	std::wstring commandLine = L"\"" + workerExePath.wstring() + L"\" --worker \"" + std::wstring(deviceId) + L"\" --stopEvent \"" + stopEventName + L"\" --appId \"" + workerAppId + L"\"";

	STARTUPINFOW startupInfo = { sizeof(startupInfo) };
	PROCESS_INFORMATION processInfo = {};
	if (!CreateProcessW(nullptr, commandLine.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startupInfo, &processInfo))
	{
		LOG_LAST_ERROR();
		CloseHandle(workerInfo.stopEventHandle);
		workerInfo.stopEventHandle = nullptr;
		return false;
	}

	CloseHandle(processInfo.hThread);
	workerInfo.processHandle = processInfo.hProcess;
	workerInfo.executablePath = workerExePath;

	if (!g_outputDeviceId.empty())
	{
		SetAppOutputDevice(workerExePath.wstring(), g_outputDeviceId);
	}

	if (g_lowLatency)
	{
		SetPriorityClass(processInfo.hProcess, HIGH_PRIORITY_CLASS);
	}

	return true;
}

void PruneExitedWorkers()
{
	for (auto it = g_workerProcesses.begin(); it != g_workerProcesses.end();)
	{
		if (WaitForSingleObject(it->second.processHandle, 0) == WAIT_OBJECT_0)
		{
			if (it->second.processHandle)
			{
				CloseHandle(it->second.processHandle);
			}
			if (it->second.stopEventHandle)
			{
				CloseHandle(it->second.stopEventHandle);
			}
			if (!it->second.executablePath.empty())
			{
				DeleteFileW(it->second.executablePath.c_str());
			}
			auto connectionIt = g_audioPlaybackConnections.find(it->first);
			if (connectionIt != g_audioPlaybackConnections.end())
			{
				try
				{
					g_devicePicker.SetDisplayStatus(connectionIt->second.first, {}, DevicePickerDisplayStatusOptions::None);
				}
				catch (winrt::hresult_error const&)
				{
					LOG_CAUGHT_EXCEPTION();
				}
				g_audioPlaybackConnections.erase(connectionIt);
			}
			it = g_workerProcesses.erase(it);
		}
		else
		{
			++it;
		}
	}
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	UNREFERENCED_PARAMETER(nCmdShow);

	std::wstring workerDeviceId;
	if (TryGetWorkerDeviceId(workerDeviceId))
	{
		std::wstring stopEventName;
		if (!TryGetArgValue(L"--stopEvent", stopEventName))
		{
			return EXIT_FAILURE;
		}
	 std::wstring workerAppId;
		if (!TryGetArgValue(L"--appId", workerAppId))
		{
			return EXIT_FAILURE;
		}
		return RunWorkerProcess(workerDeviceId, stopEventName, workerAppId);
	}

	// Prevent multiple instances
	g_hMutex = CreateMutexW(nullptr, FALSE, L"Local\\AudioPlaybackConnector_Mutex");
	if (GetLastError() == ERROR_ALREADY_EXISTS)
	{
		if (g_hMutex)
		{
			CloseHandle(g_hMutex);
			g_hMutex = nullptr;
		}
		TaskDialog(nullptr, nullptr, _(L"Already running!"), nullptr, _(L"AudioPlaybackConnector is already running in background.\r\nCheck system tray."), TDCBF_OK_BUTTON, TD_WARNING_ICON, nullptr);
		return EXIT_FAILURE;
	}

	g_hInst = hInstance;

	winrt::init_apartment();

	bool supported = false;
	try
	{
		using namespace winrt::Windows::Foundation::Metadata;

		supported = ApiInformation::IsTypePresent(winrt::name_of<DesktopWindowXamlSource>()) &&
			ApiInformation::IsTypePresent(winrt::name_of<AudioPlaybackConnection>());
	}
	catch (winrt::hresult_error const&)
	{
		supported = false;
		LOG_CAUGHT_EXCEPTION();
	}
	if (!supported)
	{
		TaskDialog(nullptr, nullptr, _(L"Unsupported Operating System"), nullptr, _(L"AudioPlaybackConnector is not supported on this operating system version."), TDCBF_OK_BUTTON, TD_ERROR_ICON, nullptr);
		return EXIT_FAILURE;
	}

	WNDCLASSEXW wcex = {
		.cbSize = sizeof(wcex),
		.lpfnWndProc = WndProc,
		.hInstance = hInstance,
		.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_AUDIOPLAYBACKCONNECTOR)),
		.hCursor = LoadCursorW(nullptr, IDC_ARROW),
		.lpszClassName = L"AudioPlaybackConnector",
		.hIconSm = wcex.hIcon
	};

	RegisterClassExW(&wcex);

	// When parent window size is 0x0 or invisible, the dpi scale of menu is incorrect. Here we set window size to 1x1 and use WS_EX_LAYERED to make window looks like invisible.
	g_hWnd = CreateWindowExW(WS_EX_NOACTIVATE | WS_EX_LAYERED | WS_EX_TOPMOST, L"AudioPlaybackConnector", nullptr, WS_POPUP, 0, 0, 0, 0, nullptr, nullptr, hInstance, nullptr);
	FAIL_FAST_LAST_ERROR_IF_NULL(g_hWnd);
	FAIL_FAST_IF_WIN32_BOOL_FALSE(SetLayeredWindowAttributes(g_hWnd, 0, 0, LWA_ALPHA));

	DesktopWindowXamlSource desktopSource;
	auto desktopSourceNative2 = desktopSource.as<IDesktopWindowXamlSourceNative2>();
	winrt::check_hresult(desktopSourceNative2->AttachToWindow(g_hWnd));
	winrt::check_hresult(desktopSourceNative2->get_WindowHandle(&g_hWndXaml));

	g_xamlCanvas = Canvas();
	desktopSource.Content(g_xamlCanvas);

	LoadTranslateData();
	LoadSettings();
	SetupFlyout();
	SetupMenu();
	SetupDevicePicker();
	SetupSvgIcon();

	g_nid.hWnd = g_niid.hWnd = g_hWnd;
	wcscpy_s(g_nid.szTip, _(L"AudioPlaybackConnector"));
	UpdateNotifyIcon();

	WM_TASKBAR_CREATED = RegisterWindowMessageW(L"TaskbarCreated");
	LOG_LAST_ERROR_IF(WM_TASKBAR_CREATED == 0);

	PostMessageW(g_hWnd, WM_CONNECTDEVICE, 0, 0);

	if (g_showNotification)
	{
		ShowInitialToastNotification();
	}

	MSG msg;
	while (GetMessageW(&msg, nullptr, 0, 0))
	{
		BOOL processed = FALSE;
		winrt::check_hresult(desktopSourceNative2->PreTranslateMessage(&msg, &processed));
		if (!processed)
		{
			TranslateMessage(&msg);
			DispatchMessageW(&msg);
		}
	}

	return static_cast<int>(msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_DESTROY:
		for (const auto& connection : g_audioPlaybackConnections)
		{
			if (connection.second.second)
			{
				connection.second.second.Close();
			}
			g_devicePicker.SetDisplayStatus(connection.second.first, {}, DevicePickerDisplayStatusOptions::None);
		}
		for (auto& worker : g_workerProcesses)
		{
			StopAndCleanupWorker(worker.second);
		}
		g_workerProcesses.clear();
		if (g_reconnect)
		{
			SaveSettings();
			g_audioPlaybackConnections.clear();
		}
		else
		{
			g_audioPlaybackConnections.clear();
			SaveSettings();
		}
		Shell_NotifyIconW(NIM_DELETE, &g_nid);
		if (g_hTrayIcon) { DestroyIcon(g_hTrayIcon); g_hTrayIcon = nullptr; }
		if (g_hMutex) { CloseHandle(g_hMutex); g_hMutex = nullptr; }
		PostQuitMessage(0);
		break;
	case WM_SETTINGCHANGE:
		if (lParam && CompareStringOrdinal(reinterpret_cast<LPCWCH>(lParam), -1, L"ImmersiveColorSet", -1, TRUE) == CSTR_EQUAL)
		{
			UpdateNotifyIcon();
		}
		break;
	case WM_NOTIFYICON:
		switch (LOWORD(lParam))
		{
		case NIN_SELECT:
		case NIN_KEYSELECT:
		{
			PruneExitedWorkers();
			using namespace winrt::Windows::UI::Popups;

			RECT iconRect;
			auto hr = Shell_NotifyIconGetRect(&g_niid, &iconRect);
			if (FAILED(hr))
			{
				LOG_HR(hr);
				break;
			}

			auto dpi = GetDpiForWindow(hWnd);
			Rect rect = {
				static_cast<float>(iconRect.left * USER_DEFAULT_SCREEN_DPI / dpi),
				static_cast<float>(iconRect.top * USER_DEFAULT_SCREEN_DPI / dpi),
				static_cast<float>((iconRect.right - iconRect.left) * USER_DEFAULT_SCREEN_DPI / dpi),
				static_cast<float>((iconRect.bottom - iconRect.top) * USER_DEFAULT_SCREEN_DPI / dpi)
			};

			SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN), SWP_HIDEWINDOW);
			SetForegroundWindow(hWnd);
			g_devicePicker.Show(rect, Placement::Above);
		}
		break;
		case WM_RBUTTONUP: // Menu activated by mouse click
			g_menuFocusState = FocusState::Pointer;
			break;
		case WM_CONTEXTMENU:
		{
			if (g_menuFocusState == FocusState::Unfocused)
				g_menuFocusState = FocusState::Keyboard;

			auto dpi = GetDpiForWindow(hWnd);
			Point point = {
				static_cast<float>(GET_X_LPARAM(wParam) * USER_DEFAULT_SCREEN_DPI / dpi),
				static_cast<float>(GET_Y_LPARAM(wParam) * USER_DEFAULT_SCREEN_DPI / dpi)
			};

			SetWindowPos(g_hWndXaml, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_SHOWWINDOW);
			SetWindowPos(g_hWnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_SHOWWINDOW);
			SetForegroundWindow(hWnd);

			g_xamlMenu.ShowAt(g_xamlCanvas, point);
		}
		break;
		}
		break;
	case WM_CONNECTIONCLOSED:
	{
		auto pDeviceId = std::unique_ptr<std::wstring>(reinterpret_cast<std::wstring*>(wParam));
		auto it = g_audioPlaybackConnections.find(*pDeviceId);
		if (it != g_audioPlaybackConnections.end())
		{
			g_devicePicker.SetDisplayStatus(it->second.first, {}, DevicePickerDisplayStatusOptions::None);
			g_audioPlaybackConnections.erase(it);
		}
		UpdateNotifyIcon();
	}
	break;
	case WM_CONNECTDEVICE:
		if (g_reconnect)
		{
			for (const auto& i : g_lastDevices)
			{
				ConnectDevice(g_devicePicker, i);
			}
			g_lastDevices.clear();
		}
		break;
	default:
		if (WM_TASKBAR_CREATED && message == WM_TASKBAR_CREATED)
		{
			UpdateNotifyIcon();
		}
		return DefWindowProcW(hWnd, message, wParam, lParam);
	}
	return 0;
}

void SetupFlyout()
{
	TextBlock textBlock;
	textBlock.Text(_(L"All connections will be closed.\nExit anyway?"));
	textBlock.Margin({ 0, 0, 0, 12 });

	static CheckBox checkbox;
	checkbox.IsChecked(g_reconnect);
	checkbox.Content(winrt::box_value(_(L"Reconnect on next start")));

	Button button;
	button.Content(winrt::box_value(_(L"Exit")));
	button.HorizontalAlignment(HorizontalAlignment::Right);
	button.Click([](const auto&, const auto&) {
		g_reconnect = checkbox.IsChecked().Value();
		PostMessageW(g_hWnd, WM_CLOSE, 0, 0);
	});

	StackPanel stackPanel;
	stackPanel.Children().Append(textBlock);
	stackPanel.Children().Append(checkbox);
	stackPanel.Children().Append(button);

	Flyout flyout;
	flyout.ShouldConstrainToRootBounds(false);
	flyout.Content(stackPanel);

	g_xamlFlyout = flyout;
}

void SetupOutputDeviceMenu(MenuFlyoutSubItem subItem)
{
	auto devices = EnumerateAudioRenderDevices();

	FontIcon checkedIcon;
	checkedIcon.Glyph(L"\xE73E");

	MenuFlyoutItem defaultItem;
	defaultItem.Text(_(L"System Default"));
	defaultItem.Tag(winrt::box_value(L"default"));
	defaultItem.Click([](const auto&, const auto&) {
		g_outputDeviceId.clear();
		SaveSettings();
		for (auto& worker : g_workerProcesses)
		{
			SetAppOutputDevice(worker.second.executablePath.wstring(), L"");
		}
	});
	subItem.Items().Append(defaultItem);

	for (const auto& dev : devices) {
		MenuFlyoutItem item;
		item.Text(dev.name);
		item.Tag(winrt::box_value(dev.id));
		item.Click([id = dev.id](const auto&, const auto&) {
			g_outputDeviceId = id;
			SaveSettings();
			for (auto& worker : g_workerProcesses)
			{
				SetAppOutputDevice(worker.second.executablePath.wstring(), g_outputDeviceId);
			}
		});
		subItem.Items().Append(item);
	}
}

void UpdateOutputDeviceMenuIcons(MenuFlyoutSubItem subItem)
{
	FontIcon checkedIcon;
	checkedIcon.Glyph(L"\xE73E");

	for (auto const& itemBase : subItem.Items())
	{
		auto item = itemBase.as<MenuFlyoutItem>();
		auto tag = winrt::unbox_value<std::wstring>(item.Tag());
		if ((g_outputDeviceId.empty() && tag == L"default") || (g_outputDeviceId == tag))
		{
			item.Icon(checkedIcon);
		}
		else
		{
			item.Icon(nullptr);
		}
	}
}

void SetupMenu()
{
	// https://docs.microsoft.com/en-us/windows/uwp/design/style/segoe-ui-symbol-font
	FontIcon settingsIcon;
	settingsIcon.Glyph(L"\xE713");

	MenuFlyoutItem settingsItem;
	settingsItem.Text(_(L"Bluetooth Settings"));
	settingsItem.Icon(settingsIcon);
	settingsItem.Click([](const auto&, const auto&) {
		winrt::Windows::System::Launcher::LaunchUriAsync(Uri(L"ms-settings:bluetooth"));
	});

	FontIcon checkedIcon, uncheckedIcon;
	checkedIcon.Glyph(L"\xE73E");

	MenuFlyoutItem startupItem;
	startupItem.Text(_(L"Run at login"));
	if (GetStartupStatus()) {
		startupItem.Icon(checkedIcon);
	}
	else {
		startupItem.Icon(uncheckedIcon);
	}
	startupItem.Click([checkedIcon, uncheckedIcon](const auto& sender, const auto&) {
		MenuFlyoutItem self = sender.as<MenuFlyoutItem>();
		if (GetStartupStatus()) {
			SetStartupStatus(false);
			self.Icon(uncheckedIcon);
		}
		else {
			SetStartupStatus(true);
			self.Icon(checkedIcon);
		}
	});

	FontIcon notificationCheckedIcon, notificationUncheckedIcon;
	notificationCheckedIcon.Glyph(L"\xE73E");

	MenuFlyoutItem notificationItem;
	notificationItem.Text(_(L"Show startup notification"));
	if (g_showNotification) {
		notificationItem.Icon(notificationCheckedIcon);
	}
	else {
		notificationItem.Icon(notificationUncheckedIcon);
	}
	notificationItem.Click([notificationCheckedIcon, notificationUncheckedIcon](const auto& sender, const auto&) {
		MenuFlyoutItem self = sender.as<MenuFlyoutItem>();
		g_showNotification = !g_showNotification;
		if (g_showNotification) {
			self.Icon(notificationCheckedIcon);
		}
		else {
			self.Icon(notificationUncheckedIcon);
		}
		SaveSettings();
	});

	MenuFlyoutItem lowLatencyItem;
	lowLatencyItem.Text(_(L"Low Latency Mode"));
	if (g_lowLatency) {
		lowLatencyItem.Icon(checkedIcon);
	}
	else {
		lowLatencyItem.Icon(uncheckedIcon);
	}
	lowLatencyItem.Click([checkedIcon, uncheckedIcon](const auto& sender, const auto&) {
		MenuFlyoutItem self = sender.as<MenuFlyoutItem>();
		g_lowLatency = !g_lowLatency;
		if (g_lowLatency) {
			self.Icon(checkedIcon);
		}
		else {
			self.Icon(uncheckedIcon);
		}
		SaveSettings();
		for (auto& worker : g_workerProcesses)
		{
			SetPriorityClass(worker.second.processHandle, g_lowLatency ? HIGH_PRIORITY_CLASS : NORMAL_PRIORITY_CLASS);
		}
	});

	FontIcon speakerIcon;
	speakerIcon.Glyph(L"\xE767");

	MenuFlyoutSubItem outputDeviceItem;
	outputDeviceItem.Text(_(L"Output Device"));
	outputDeviceItem.Icon(speakerIcon);
	SetupOutputDeviceMenu(outputDeviceItem);

	MenuFlyoutItem soundSettingsItem;
	soundSettingsItem.Text(_(L"Open Sound Settings"));
	FontIcon volumeIcon;
	volumeIcon.Glyph(L"\xE767");
	soundSettingsItem.Icon(volumeIcon);
	soundSettingsItem.Click([](const auto&, const auto&) {
		winrt::Windows::System::Launcher::LaunchUriAsync(Uri(L"ms-settings:sound"));
	});

	FontIcon closeIcon;
	closeIcon.Glyph(L"\xE8BB");

	MenuFlyoutItem exitItem;
	exitItem.Text(_(L"Exit"));
	exitItem.Icon(closeIcon);
	exitItem.Click([](const auto&, const auto&) {
		if (g_audioPlaybackConnections.size() == 0)
		{
			PostMessageW(g_hWnd, WM_CLOSE, 0, 0);
			return;
		}

		RECT iconRect;
		auto hr = Shell_NotifyIconGetRect(&g_niid, &iconRect);
		if (FAILED(hr))
		{
			LOG_HR(hr);
			return;
		}

		auto dpi = GetDpiForWindow(g_hWnd);

		SetWindowPos(g_hWnd, HWND_TOPMOST, iconRect.left, iconRect.top, 0, 0, SWP_HIDEWINDOW);
		g_xamlCanvas.Width(static_cast<float>((iconRect.right - iconRect.left) * USER_DEFAULT_SCREEN_DPI / dpi));
		g_xamlCanvas.Height(static_cast<float>((iconRect.bottom - iconRect.top) * USER_DEFAULT_SCREEN_DPI / dpi));

		g_xamlFlyout.ShowAt(g_xamlCanvas);
	});

	MenuFlyout menu;
	menu.Items().Append(settingsItem);
	menu.Items().Append(soundSettingsItem);
	menu.Items().Append(outputDeviceItem);
	menu.Items().Append(startupItem);
	menu.Items().Append(notificationItem);
	menu.Items().Append(lowLatencyItem);
	menu.Items().Append(exitItem);
	menu.Opened([outputDeviceItem](const auto& sender, const auto&) {
		UpdateOutputDeviceMenuIcons(outputDeviceItem);
		auto menuItems = sender.as<MenuFlyout>().Items();
		auto itemsCount = menuItems.Size();
		if (itemsCount > 0)
		{
			menuItems.GetAt(itemsCount - 1).Focus(g_menuFocusState);
		}
		g_menuFocusState = FocusState::Unfocused;
	});
	menu.Closed([](const auto&, const auto&) {
		ShowWindow(g_hWnd, SW_HIDE);
	});

	g_xamlMenu = menu;
}

winrt::fire_and_forget ConnectDevice(DevicePicker picker, DeviceInformation device)
{
	try
	{
		PruneExitedWorkers();

		auto deviceId = std::wstring(device.Id());
		auto existing = g_audioPlaybackConnections.find(deviceId);
		if (existing != g_audioPlaybackConnections.end())
		{
			picker.SetDisplayStatus(device, _(L"Connected"), DevicePickerDisplayStatusOptions::ShowDisconnectButton);
			co_return;
		}

		picker.SetDisplayStatus(device, _(L"Connecting"), DevicePickerDisplayStatusOptions::ShowProgress | DevicePickerDisplayStatusOptions::ShowDisconnectButton);

		WorkerProcessInfo workerInfo;
		if (!LaunchWorkerProcess(deviceId, workerInfo))
		{
			picker.SetDisplayStatus(device, _(L"Unknown error"), DevicePickerDisplayStatusOptions::ShowRetryButton);
			co_return;
		}

		g_workerProcesses.emplace(deviceId, workerInfo);
		g_audioPlaybackConnections.emplace(deviceId, std::pair(device, AudioPlaybackConnection{ nullptr }));
		picker.SetDisplayStatus(device, _(L"Connected"), DevicePickerDisplayStatusOptions::ShowDisconnectButton);
		UpdateNotifyIcon();
	}
	catch (winrt::hresult_error const&)
	{
		LOG_CAUGHT_EXCEPTION();
	}
}

winrt::fire_and_forget ConnectDevice(DevicePicker picker, std::wstring deviceId)
{
	auto device = co_await DeviceInformation::CreateFromIdAsync(deviceId);
	ConnectDevice(picker, device);
}

void SetupDevicePicker()
{
	g_devicePicker = DevicePicker();
	winrt::check_hresult(g_devicePicker.as<IInitializeWithWindow>()->Initialize(g_hWnd));

	g_devicePicker.Filter().SupportedDeviceSelectors().Append(AudioPlaybackConnection::GetDeviceSelector());
	g_devicePicker.DevicePickerDismissed([](const auto&, const auto&) {
		SetWindowPos(g_hWnd, nullptr, 0, 0, 0, 0, SWP_NOZORDER | SWP_HIDEWINDOW);
	});
	g_devicePicker.DeviceSelected([](const auto& sender, const auto& args) {
		ConnectDevice(sender, args.SelectedDevice());
	});
	g_devicePicker.DisconnectButtonClicked([](const auto& sender, const auto& args) {
		auto device = args.Device();
		auto deviceId = std::wstring(device.Id());

		auto it = g_workerProcesses.find(deviceId);
		if (it != g_workerProcesses.end())
		{
			StopAndCleanupWorker(it->second);
			g_workerProcesses.erase(it);
		}

		auto itConn = g_audioPlaybackConnections.find(deviceId);
		if (itConn != g_audioPlaybackConnections.end())
		{
			sender.SetDisplayStatus(itConn->second.first, {}, DevicePickerDisplayStatusOptions::None);
			g_audioPlaybackConnections.erase(itConn);
		}

		UpdateNotifyIcon();
	});
}

void SetupSvgIcon()
{
	auto hRes = FindResourceW(g_hInst, MAKEINTRESOURCEW(1), L"SVG");
	FAIL_FAST_LAST_ERROR_IF_NULL(hRes);

	auto size = SizeofResource(g_hInst, hRes);
	FAIL_FAST_LAST_ERROR_IF(size == 0);

	auto hResData = LoadResource(g_hInst, hRes);
	FAIL_FAST_LAST_ERROR_IF_NULL(hResData);

	auto svgData = reinterpret_cast<const char*>(LockResource(hResData));
	FAIL_FAST_IF_NULL_ALLOC(svgData);

	g_notifyIconSvg.assign(svgData, size);
}

void UpdateNotifyIcon()
{
	PruneExitedWorkers();

	auto icon = CreateNotifyIcon(g_audioPlaybackConnections.size());
	if (icon)
	{
		if (g_hTrayIcon)
		{
			DestroyIcon(g_hTrayIcon);
		}
		g_hTrayIcon = icon;
		g_nid.hIcon = g_hTrayIcon;
	}

	if (!Shell_NotifyIconW(NIM_MODIFY, &g_nid))
	{
		if (Shell_NotifyIconW(NIM_ADD, &g_nid))
		{
			FAIL_FAST_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_SETVERSION, &g_nid));
		}
		else
		{
			LOG_LAST_ERROR();
		}
	}
}

bool GetStartupStatus()
{
	auto exePath = GetModuleFsPath(g_hInst);

	wil::unique_hkey hKey;
	if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
	{
		wchar_t storedPath[MAX_PATH] = { 0 };
		DWORD pathLength = sizeof(storedPath);
		DWORD type = REG_SZ;
		LSTATUS result = RegQueryValueExW(hKey.get(), L"AudioPlaybackConnector", 0, &type, (LPBYTE)storedPath, &pathLength);

		if (result == ERROR_SUCCESS && type == REG_SZ && exePath == storedPath)
		{
			return true;
		}
	}

	return false;
}

void SetStartupStatus(bool status)
{
	auto exePath = GetModuleFsPath(g_hInst);

	wil::unique_hkey hKey;
	if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
	{
		if (status)
		{
			auto exePathStr = exePath.wstring();
			LOG_IF_WIN32_ERROR(RegSetValueExW(hKey.get(), L"AudioPlaybackConnector", 0, REG_SZ, (LPBYTE)exePathStr.c_str(), (lstrlenW(exePathStr.c_str()) + 1) * sizeof(wchar_t)));
		}
		else
		{
			LOG_IF_WIN32_ERROR(RegDeleteValueW(hKey.get(), L"AudioPlaybackConnector"));
		}
	}
}

void ShowInitialToastNotification()
{
	try
	{
		std::wstring title = _(L"AudioPlaybackConnector");
		std::wstring message = _(L"Application has started and is running in the notification area.");

		std::wstring toastXmlString =
			L"<toast activationType=\"protocol\" launch=\"audioplaybackconnector:show\">" 
			L"<visual>"
			L"<binding template=\"ToastGeneric\">"
			L"<text>" + title + L"</text>"
			L"<text>" + message + L"</text>"
			L"</binding>"
			L"</visual>"
			L"</toast>";

		XmlDocument toastXml;
		toastXml.LoadXml(toastXmlString);

		ToastNotifier notifier{ nullptr };
		try
		{
			notifier = ToastNotificationManager::CreateToastNotifier();
		}
		catch (winrt::hresult_error const&)
		{
			LOG_CAUGHT_EXCEPTION();
			// wchar_t exePath[MAX_PATH];
			// GetModuleFileNameW(NULL, exePath, MAX_PATH);
			// std::wstring appId = exePath;
			try
			{
				// notifier = ToastNotificationManager::CreateToastNotifier(appId);
				notifier = ToastNotificationManager::CreateToastNotifier(L"AudioPlaybackConnector");
			}
			catch (winrt::hresult_error const&)
			{
				LOG_CAUGHT_EXCEPTION();
				return;
			}
		}

		// if (!notifier)
		// {
		// 	return;
		// }

		ToastNotification toast(toastXml);

		toast.ExpirationTime(winrt::Windows::Foundation::DateTime::clock::now() + std::chrono::seconds(5));

		notifier.Show(toast);
	}
	catch (winrt::hresult_error const&)
	{
		LOG_CAUGHT_EXCEPTION();
	}
	catch (std::exception const&)
	{
		// Silently ignore standard exceptions from toast notification - this is not critical functionality
	}
}

fs::path GetWorkerExecutablePath(std::wstring_view deviceId, uint64_t launchId)
{
	auto baseDir = GetModuleFsPath(g_hInst).remove_filename();
	auto workersDir = baseDir / L"workers";
	std::wstring id(deviceId);
	auto hash = fnv1a_32(id.data(), id.size() * sizeof(wchar_t));
	wchar_t fileName[96] = {};
	swprintf_s(fileName, L"AudioPlaybackConnectorWorker_%08X_%llu.exe", hash, static_cast<unsigned long long>(launchId));
	return workersDir / fileName;
}

bool EnsureWorkerExecutable(const fs::path& workerExePath)
{
	try
	{
		fs::create_directories(workerExePath.parent_path());

		auto sourceExe = GetModuleFsPath(g_hInst);
		if (!CopyFileW(sourceExe.c_str(), workerExePath.c_str(), FALSE))
		{
			LOG_LAST_ERROR();
			return false;
		}
		return true;
	}
	catch (...)
	{
		LOG_CAUGHT_EXCEPTION();
		return false;
	}
}

std::wstring GetWorkerAppId(std::wstring_view deviceId, uint64_t launchId)
{
	std::wstring id(deviceId);
	auto hash = fnv1a_32(id.data(), id.size() * sizeof(wchar_t));
	wchar_t appId[128] = {};
	swprintf_s(appId, L"AudioPlaybackConnector.Worker.%08X.%llu", hash, static_cast<unsigned long long>(launchId));
	return appId;
}
