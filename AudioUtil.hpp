#pragma once
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>

// Undocumented IPolicyConfig10 interface
MIDL_INTERFACE("CA286FC3-91FD-42C3-8E9B-CAAFA66242E3")
IPolicyConfig10 : public IUnknown {
public:
    virtual HRESULT STDMETHODCALLTYPE GetMixFormat(PCWSTR, WAVEFORMATEX**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDeviceFormat(PCWSTR, INT, WAVEFORMATEX**) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetDeviceFormat(PCWSTR, WAVEFORMATEX*, WAVEFORMATEX*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetProcessingPeriod(PCWSTR, INT, REFERENCE_TIME*, REFERENCE_TIME*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetProcessingPeriod(PCWSTR, REFERENCE_TIME*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetShareMode(PCWSTR, DeviceShareMode*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetShareMode(PCWSTR, DeviceShareMode) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(PCWSTR, const PROPERTYKEY&, PROPVARIANT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(PCWSTR, const PROPERTYKEY&, PROPVARIANT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(PCWSTR wszDeviceId, ERole eRole) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetEndpointVisibility(PCWSTR wszDeviceId, INT isVisible) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetAppDefaultEndpoint(PCWSTR wszAppPath, ERole eRole, PCWSTR wszDeviceId) = 0;
};

inline HRESULT SetAppOutputDevice(const std::wstring& appPath, const std::wstring& deviceId)
{
    if (deviceId.empty()) return S_OK;

    wil::com_ptr<IPolicyConfig10> policyConfig;
    HRESULT hr = CoCreateInstance(__uuidof(CPolicyConfigClient), nullptr, CLSCTX_ALL, __uuidof(IPolicyConfig10), policyConfig.put_void());
    if (FAILED(hr)) return hr;

    hr = policyConfig->SetAppDefaultEndpoint(appPath.c_str(), eConsole, deviceId.c_str());
    if (FAILED(hr)) return hr;

    hr = policyConfig->SetAppDefaultEndpoint(appPath.c_str(), eMultimedia, deviceId.c_str());
    return hr;
}

struct AudioDeviceDetails {
    std::wstring id;
    std::wstring name;
};

inline std::vector<AudioDeviceDetails> EnumerateAudioRenderDevices()
{
    std::vector<AudioDeviceDetails> devices;
    try {
        wil::com_ptr<IMMDeviceEnumerator> enumerator;
        THROW_IF_FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), enumerator.put_void()));

        wil::com_ptr<IMMDeviceCollection> collection;
        THROW_IF_FAILED(enumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &collection));

        UINT count = 0;
        THROW_IF_FAILED(collection->GetCount(&count));

        for (UINT i = 0; i < count; ++i) {
            wil::com_ptr<IMMDevice> device;
            THROW_IF_FAILED(collection->Item(i, &device));

            wil::unique_cotaskmem_string id;
            THROW_IF_FAILED(device->GetId(&id));

            wil::com_ptr<IPropertyStore> properties;
            THROW_IF_FAILED(device->OpenPropertyStore(STGM_READ, &properties));

            wil::unique_prop_variant nameVar;
            THROW_IF_FAILED(properties->GetValue(PKEY_Device_FriendlyName, &nameVar));

            devices.push_back({ id.get(), nameVar.pwszVal });
        }
    }
    CATCH_LOG();
    return devices;
}
