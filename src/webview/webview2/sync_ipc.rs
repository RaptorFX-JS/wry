// Copyright 2022 The RaptorFX Team, ReMod Software
// SPDX-License-Identifier: Apache-2.0
// SPDX-License-Identifier: MIT

use std::{
  mem::{size_of, ManuallyDrop},
  rc::Rc,
};

use windows::{
  core::{IUnknown, IUnknownVtbl, Interface, GUID, PCWSTR, PWSTR},
  Win32::{
    Foundation::{BSTR, DISP_E_BADINDEX, DISP_E_UNKNOWNINTERFACE},
    Globalization::{LocaleNameToLCID, LOCALE_NAME_INVARIANT},
    System::{
      Com::{
        IDispatch, IDispatch_Impl, ITypeInfo, CC_STDCALL, DISPPARAMS, EXCEPINFO, VARIANT,
        VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
      },
      Ole::{
        CreateDispTypeInfo, DispGetIDsOfNames, DispInvoke, DISPATCH_METHOD, INTERFACEDATA,
        METHODDATA, PARAMDATA, VT_BSTR, VT_DISPATCH,
      },
    },
  },
};

use windows_implement::implement;
use windows_interface::interface;

use webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2;

use crate::application::window::Window;

macro_rules! pwstr {
  ($string:literal) => {{
    const UTF16: &[u16] = ::const_utf16::encode_null_terminated!($string);
    static mut MUTABLE_UTF16: [u16; UTF16.len()] = {
      let mut out = [0; UTF16.len()];
      unsafe {
        ::std::ptr::copy_nonoverlapping(UTF16.as_ptr(), out.as_mut_ptr(), UTF16.len());
      }
      out
    };
    unsafe { ::windows::core::PWSTR(&mut MUTABLE_UTF16 as *mut _) }
  }};
}

#[interface("e0912f1d-f683-40cd-94c6-20a1d7e96bdc")]
unsafe trait ISyncIPCHandler: IUnknown {
  unsafe fn PostSyncMessage(&self, message: BSTR) -> BSTR;
}

#[implement(IDispatch, ISyncIPCHandler)]
pub(crate) struct SyncIPCHandler {
  type_info: ITypeInfo,
  window: Rc<Window>,
  handler: Box<dyn Fn(&Window, String) -> String>,
}

impl SyncIPCHandler {
  pub(crate) fn new(
    window: Rc<Window>,
    handler: Box<dyn Fn(&Window, String) -> String>,
  ) -> webview2_com::Result<Self> {
    // Safety: we never mutate SyncIPCHandler's type information, so the below statics can be safely Sync
    #[repr(transparent)]
    struct SyncStatic<T>(T);

    unsafe impl<T> Sync for SyncStatic<T> {}

    static mut DISPATCH_INTERFACE_POST_SYNC_MESSAGE_PARAMS: SyncStatic<[PARAMDATA; 1]> =
      SyncStatic([PARAMDATA {
        szName: pwstr!("message"),
        vt: VT_BSTR.0 as u16,
      }]);

    static mut DISPATCH_INTERFACE_METHODS: SyncStatic<[METHODDATA; 1]> = SyncStatic([METHODDATA {
      szName: pwstr!("PostSyncMessage"),
      ppdata: unsafe { &mut DISPATCH_INTERFACE_POST_SYNC_MESSAGE_PARAMS.0 as *mut _ },
      dispid: 0,
      // PostSyncMessage is the first method in ISyncIPCHandler
      #[allow(clippy::identity_op)]
      iMeth: (size_of::<IUnknownVtbl>() / size_of::<fn()>() + 0) as u32,
      cc: CC_STDCALL,
      cArgs: unsafe { DISPATCH_INTERFACE_METHODS.0.len() as u32 },
      wFlags: DISPATCH_METHOD as u16,
      vtReturn: VT_BSTR.0 as u16,
    }]);

    static mut DISPATCH_INTERFACE: SyncStatic<INTERFACEDATA> = SyncStatic(INTERFACEDATA {
      pmethdata: unsafe { &mut DISPATCH_INTERFACE_METHODS.0 as *mut _ },
      cMembers: unsafe { DISPATCH_INTERFACE_METHODS.0.len() as u32 },
    });

    // Safety: WinAPI calls are unsafe
    let type_info = unsafe {
      let invariant_locale = LocaleNameToLCID(LOCALE_NAME_INVARIANT, 0);
      let mut type_info = None;
      CreateDispTypeInfo(
        &mut DISPATCH_INTERFACE.0 as *mut _,
        invariant_locale,
        &mut type_info as *mut _,
      )?;
      type_info.unwrap()
    };

    Ok(Self {
      type_info,
      window,
      handler,
    })
  }

  pub(crate) fn inject(self, webview: &ICoreWebView2) -> webview2_com::Result<()> {
    let handler: IDispatch = self.into();

    // Safety: WinAPI calls are unsafe
    unsafe {
      // wrapper struct to ensure that VARIANT ManuallyDrops are actually dropped
      #[repr(transparent)]
      struct IDispatchVariant(VARIANT);

      impl Drop for IDispatchVariant {
        fn drop(&mut self) {
          unsafe {
            if self.0.Anonymous.Anonymous.vt == VT_DISPATCH.0 as u16 {
              ManuallyDrop::drop(&mut (&mut self.0.Anonymous.Anonymous).Anonymous.pdispVal);
            }
            ManuallyDrop::drop(&mut self.0.Anonymous.Anonymous);
          }
        }
      }

      let mut remote_obj = IDispatchVariant(VARIANT {
        Anonymous: VARIANT_0 {
          Anonymous: ManuallyDrop::new(VARIANT_0_0 {
            vt: VT_DISPATCH.0 as u16,
            Anonymous: VARIANT_0_0_0 {
              pdispVal: ManuallyDrop::new(Some(handler)),
            },
            ..VARIANT_0_0::default()
          }),
        },
      });

      webview
        .AddHostObjectToScript(
          PCWSTR(const_utf16::encode_null_terminated!("ipc") as *const _),
          &mut remote_obj.0 as *mut _,
        )
        .map_err(webview2_com::Error::WindowsError)
    }
  }
}

#[allow(non_snake_case)]
impl IDispatch_Impl for SyncIPCHandler {
  fn GetTypeInfoCount(&self) -> windows::core::Result<u32> {
    Ok(1)
  }

  fn GetTypeInfo(&self, itinfo: u32, _lcid: u32) -> windows::core::Result<ITypeInfo> {
    if itinfo != 0 {
      Err(DISP_E_BADINDEX.into())
    } else {
      Ok(self.type_info.clone())
    }
  }

  fn GetIDsOfNames(
    &self,
    riid: *const GUID,
    rgsznames: *const PWSTR,
    cnames: u32,
    _lcid: u32,
    rgdispid: *mut i32,
  ) -> windows::core::Result<()> {
    // Safety: riid is checked for null before deref + WinAPI calls are unsafe
    unsafe {
      if riid.is_null() || *riid != GUID::default() {
        Err(DISP_E_UNKNOWNINTERFACE.into())
      } else {
        DispGetIDsOfNames(&self.type_info, rgsznames, cnames, rgdispid)
      }
    }
  }

  fn Invoke(
    &self,
    dispidmember: i32,
    riid: *const GUID,
    _lcid: u32,
    wflags: u16,
    pdispparams: *const DISPPARAMS,
    pvarresult: *mut VARIANT,
    pexcepinfo: *mut EXCEPINFO,
    puargerr: *mut u32,
  ) -> windows::core::Result<()> {
    // Safety: pointers are checked for null before deref + WinAPI calls are unsafe
    unsafe {
      if riid.is_null() || *riid != GUID::default() {
        Err(DISP_E_UNKNOWNINTERFACE.into())
      } else {
        let this: ISyncIPCHandler = self.cast()?;

        // Invoke takes a *const DISPPARAMS but DispInvoke wants a *mut DISPPARAMS ???
        let mut dispparams = if pdispparams.is_null() {
          None
        } else {
          Some(*pdispparams)
        };
        let pdispparams_mut = dispparams
          .as_mut()
          .map(|x| x as _)
          .unwrap_or(std::ptr::null_mut());

        DispInvoke(
          this.as_raw(),
          &self.type_info,
          dispidmember,
          wflags,
          pdispparams_mut,
          pvarresult,
          pexcepinfo,
          puargerr,
        )
      }
    }
  }
}

#[allow(non_snake_case)]
impl ISyncIPCHandler_Impl for SyncIPCHandler {
  unsafe fn PostSyncMessage(&self, message: BSTR) -> BSTR {
    if let Ok(utf8_message) = message.try_into() {
      (self.handler)(&self.window, utf8_message).into()
    } else {
      BSTR::default()
    }
  }
}
