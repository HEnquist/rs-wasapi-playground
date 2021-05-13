fn main() {
    windows::build!(
        Windows::Win32::Media::Audio::CoreAudio::*,
        Windows::Win32::Devices::FunctionDiscovery::IFunctionInstance,
        Windows::Win32::Media::Multimedia::{
            WAVEFORMATEX,
            WAVEFORMATEXTENSIBLE,
            WAVE_FORMAT_PCM,
            WAVE_FORMAT_IEEE_FLOAT,
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
        },
        Windows::Win32::Media::Audio::DirectMusic::IPropertyStore,
        Windows::Win32::System::Com::{COINIT_MULTITHREADED, CoTaskMemAlloc, CoTaskMemFree, CLSIDFromProgID, CoInitializeEx, CoCreateInstance, CLSCTX},
        Windows::Win32::System::Threading::{
            CreateEventA,
            ResetEvent,
            SetEvent,
            WAIT_RETURN_CAUSE,
            WaitForSingleObject,
            WaitForMultipleObjects,
        },
        Windows::Win32::System::SystemServices::{
            HANDLE,
            INVALID_HANDLE_VALUE,
            FALSE,
            TRUE,
            S_FALSE,
        },
        Windows::Win32::System::PropertiesSystem::PROPERTYKEY,
        Windows::Win32::System::SystemServices::PWSTR,
        Windows::Win32::Storage::StructuredStorage::{STGM_READ, PROPVARIANT},
        Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
        Windows::Win32::System::WindowsProgramming::{INFINITE, CloseHandle},
        Windows::Win32::System::ApplicationInstallationAndServicing::NTDDI_WIN7,
    );
}