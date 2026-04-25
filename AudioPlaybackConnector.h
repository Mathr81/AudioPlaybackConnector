#pragma once

#include "resource.h"

namespace fs = std::filesystem;

constexpr UINT WM_NOTIFYICON = WM_APP + 1;
constexpr UINT WM_CONNECTDEVICE = WM_APP + 2;
constexpr UINT WM_CONNECTIONCLOSED = WM_APP + 3;

HANDLE g_hMutex = nullptr;
HINSTANCE g_hInst;
HWND g_hWnd;
HWND g_hWndXaml;
winrt::Windows::UI::Xaml::Controls::Canvas g_xamlCanvas = nullptr;
winrt::Windows::UI::Xaml::Controls::Flyout g_xamlFlyout = nullptr;
winrt::Windows::UI::Xaml::Controls::MenuFlyout g_xamlMenu = nullptr;
winrt::Windows::UI::Xaml::FocusState g_menuFocusState = winrt::Windows::UI::Xaml::FocusState::Unfocused;
winrt::Windows::Devices::Enumeration::DevicePicker g_devicePicker = nullptr;
struct WorkerProcessInfo
{
	HANDLE processHandle = nullptr;
	HANDLE stopEventHandle = nullptr;
	fs::path executablePath;
};
std::unordered_map<std::wstring, std::pair<winrt::Windows::Devices::Enumeration::DeviceInformation, winrt::Windows::Media::Audio::AudioPlaybackConnection>> g_audioPlaybackConnections;
std::unordered_map<std::wstring, WorkerProcessInfo> g_workerProcesses;
uint64_t g_workerEventSerial = 0;
std::string g_notifyIconSvg;
HICON g_hTrayIcon = nullptr;
NOTIFYICONDATAW g_nid = {
	.cbSize = sizeof(g_nid),
	.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP,
	.uCallbackMessage = WM_NOTIFYICON,
	.uVersion = NOTIFYICON_VERSION_4
};
NOTIFYICONIDENTIFIER g_niid = {
	.cbSize = sizeof(g_niid)
};
UINT WM_TASKBAR_CREATED = 0;
bool g_reconnect = false;
bool g_showNotification = true;
bool g_lowLatency = false;
std::vector<std::wstring> g_lastDevices;
std::wstring g_outputDeviceId;

#include "Util.hpp"
#include "FnvHash.hpp"
#include "I18n.hpp"
#include "SettingsUtil.hpp"
#include "Direct2DSvg.hpp"
