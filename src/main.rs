use wasapi::{
    Windows::Win32::Media::Audio::CoreAudio::{IMMDevice, eRender, eConsole, MMDeviceEnumerator, IMMDeviceEnumerator, DEVICE_STATE_ACTIVE},
    Windows::Win32::System::PropertiesSystem::PROPERTYKEY,
    Windows::Win32::Storage::StructuredStorage::PROPVARIANT,
    Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    Windows::Win32::System::SystemServices::PWSTR,
    Windows::Win32::Storage::StructuredStorage::STGM_READ,
    Windows::Win32::System::Com::{COINIT_MULTITHREADED, CoTaskMemAlloc, CoTaskMemFree, CLSIDFromProgID, CoInitializeEx},
    PKEY_Device_FriendlyName,
};
use std::mem;
use widestring::U16CString;

fn main() -> windows::Result<()> {
    unsafe  {
        let res = CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED);
        println!("{:?}", res);
    }
    let mut device = None;
    let enumerator: IMMDeviceEnumerator = windows::create_instance(&MMDeviceEnumerator)?;
    unsafe {
        enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole, &mut device)
            .ok()?;
        println!("{:?}", device);

        if let Some(device) = device {
            println!("{:?}", device);
            let mut store = None;
            device.OpenPropertyStore(STGM_READ as u32, &mut store).ok()?;
            let mut state: u32 = 0;
            device.GetState(&mut state).ok()?;
            println!("{:?}", state);
            let mut prop: mem::MaybeUninit<PROPVARIANT> = mem::MaybeUninit::zeroed();
            let mut propstr = PWSTR::NULL;
            println!("read prop into {:?}", propstr);
            store.unwrap().GetValue(&PKEY_Device_FriendlyName, prop.as_mut_ptr() as *mut _).ok()?;
            let prop = prop.assume_init();
            println!("read prop");
            PropVariantToStringAlloc(&prop, &mut propstr).ok()?;
            // Get the buffer as a wide string
            let wide_string = U16CString::from_ptr_str(propstr.0);
            //let name = PWSTR::try_from(prop); 
            println!("{}", wide_string.to_string_lossy());
        }
    }
    unsafe {
        let mut devs = None;
        enumerator
            .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &mut devs)
            .ok()?;
        let devs = devs.unwrap();
        println!("{:?}", devs);
        let mut count = 0;
        devs.GetCount(&mut count).ok()?;
        println!("count {}", count);
        for n in 0..count {
            let mut device = None;
            devs.Item(n, &mut device).unwrap();
            let device = device.unwrap();
            let mut store = None;
            device.OpenPropertyStore(STGM_READ as u32, &mut store).ok()?;
            let mut state: u32 = 0;
            device.GetState(&mut state).ok()?;
            println!("{:?}", state);
            let mut prop: mem::MaybeUninit<PROPVARIANT> = mem::MaybeUninit::zeroed();
            let mut propstr = PWSTR::NULL;
            println!("read prop into {:?}", propstr);
            store.unwrap().GetValue(&PKEY_Device_FriendlyName, prop.as_mut_ptr() as *mut _).ok()?;
            let prop = prop.assume_init();
            println!("read prop");
            PropVariantToStringAlloc(&prop, &mut propstr).ok()?;
            let wide_string = U16CString::from_ptr_str(propstr.0);
            //let name = PWSTR::try_from(prop); 
            println!("{}", wide_string.to_string_lossy());
        }
    }
    println!("done");
    Ok(())
}