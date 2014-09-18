/*!

Shows the "Run" dialog under Windows using COM automation.

*/
#![allow(dead_code)]
#![feature(macro_rules)]
#![license = "MIT"]

extern crate libc;

macro_rules! DEFINE_GUID(
	($name:ident, $l:expr, $w1:expr, $w2:expr, $($bs:expr),+) => {
		pub static $name: ::types::GUID = ::types::GUID {
			data1: $l,
			data2: $w1,
			data3: $w2,
			data4: [$($bs),+]
		};
	};
)

fn main() {
	/*
	Until `rustc` stops using the LLVM segmented stacks support on Windows, you have to call this to un-bork OS-private data.

	In this particular case, it causes CoCreateInstance to fail with ERROR_NOACCESS.
	*/
	//fix_corrupt_tlb();
	win32::show_run_file_dialog();
}

fn fix_corrupt_tlb() {
	unsafe { ::std::rt::stack::record_sp_limit(0); }
}

mod win32 {
	use libc::c_void;
	use types::{DWORD, HRESULT, REFCLSID, REFIID};
	use types::{IUnknown};

	pub fn show_run_file_dialog() {
		use std::mem::transmute;
		use std::ptr::mut_null;
		use types::{IID_IShellDispatch, IShellDispatch};

		match unsafe { CoInitializeEx(mut_null(), COINIT_APARTMENTTHREADED) } {
			S_OK => (),
			S_FALSE => (),
			result => fail!("call to CoInitializeEx failed: {}", result)
		}

		/*
		You know what would make this easier?  Being able to do this:
		
			match CoCreateInstance(..., &mut let obj) { ... }

			// obj is in scope here

		Not *widely* useful, especially not in idiomatic Rust, but still nice for dealing with these sorts of APIs.
		*/
		let mut obj: *mut IShellDispatch = mut_null();
		match unsafe { CoCreateInstance(&CLSID_ShApp, mut_null(), CLSCTX_INPROC_SERVER, &IID_IShellDispatch, transmute(&mut obj)) } {
			S_OK => (),
			REGDB_E_CLASSNOTREG => fail!("CoCreateInstance failed: class not registered"),
			E_NOINTERFACE => fail!("CoCreateInstance failed: class does not implement interface"),
			0x800703e6 => fail!("CoCreateInstance failed: ERROR_NOACCESS; see https://github.com/rust-lang/rust/issues/13259"),
			result => fail!("CoCreateInstance failed: error {:#08x}", result)
		}

		assert!(obj != mut_null());

		match unsafe { ((*(*obj).__vtable).FileRun)(transmute(obj)) } {
			S_OK => (),
			result => fail!("IShellDispatch.FileRun failed: error {:#08x}", result)
		}

		unsafe {
			((*(*obj).__vtable).__base.__base.Release)(transmute(obj));
		}

		unsafe {
			CoUninitialize();
		}
	}

	#[link(name = "ole32")]
	extern "stdcall" {
		fn CoCreateInstance(rclsid: REFCLSID, pUnkOuter: *mut IUnknown, dwClsContext: DWORD, riid: REFIID, ppv: *mut *mut c_void) -> HRESULT;
		fn CoInitializeEx(pvReserved: *mut c_void, dwCoInit: DWORD) -> HRESULT;
		fn CoUninitialize();
	}

	static CLSCTX_INPROC_SERVER: DWORD = 0x1;
	static CLSCTX_INPROC_HANDLER: DWORD = 0x2;
	static CLSCTX_LOCAL_SERVER: DWORD = 0x4;
	static CLSCTX_INPROC_SERVER16: DWORD = 0x8;
	static CLSCTX_REMOTE_SERVER: DWORD = 0x10;
	static CLSCTX_INPROC_HANDLER16: DWORD = 0x20;
	static CLSCTX_RESERVED1: DWORD = 0x40;
	static CLSCTX_RESERVED2: DWORD = 0x80;
	static CLSCTX_RESERVED3: DWORD = 0x100;
	static CLSCTX_RESERVED4: DWORD = 0x200;
	static CLSCTX_NO_CODE_DOWNLOAD: DWORD = 0x400;
	static CLSCTX_RESERVED5: DWORD = 0x800;
	static CLSCTX_NO_CUSTOM_MARSHAL: DWORD = 0x1000;
	static CLSCTX_ENABLE_CODE_DOWNLOAD: DWORD = 0x2000;
	static CLSCTX_NO_FAILURE_LOG: DWORD = 0x4000;
	static CLSCTX_DISABLE_AAA: DWORD = 0x8000;
	static CLSCTX_ENABLE_AAA: DWORD = 0x10000;
	static CLSCTX_FROM_DEFAULT_CONTEXT: DWORD = 0x20000;
	static CLSCTX_ACTIVATE_32_BIT_SERVER: DWORD = 0x40000;
	static CLSCTX_ACTIVATE_64_BIT_SERVER: DWORD = 0x80000;
	static CLSCTX_ENABLE_CLOAKING: DWORD = 0x100000;
	static CLSCTX_APPCONTAINER: DWORD = 0x400000;
	static CLSCTX_ACTIVATE_AAA_AS_IU: DWORD = 0x800000;
	static CLSCTX_PS_DLL: DWORD = 0x80000000;

	static COINIT_APARTMENTTHREADED: DWORD = 0x2;
	static COINIT_MULTITHREADED: DWORD = 0x0;
	static COINIT_DISABLE_OLE1DDE: DWORD = 0x4;
	static COINIT_SPEED_OVER_MEMORY: DWORD = 0x8;

	static S_OK: HRESULT = 0x00000000;
	static S_FALSE: HRESULT = 0x00000001;

	static CLASS_E_NOAGGREGATION: HRESULT = 0x80040110;
	static E_NOINTERFACE: HRESULT = 0x80004002;
	static REGDB_E_CLASSNOTREG: HRESULT = 0x80040154;
	static REGDB_E_IIDNOTREG: HRESULT = 0x80040155;
	static RPC_E_CHANGED_MODE: HRESULT = 0x80010106;

	// {13709620-C279-11CE-A49E-444553540000}
	DEFINE_GUID!(CLSID_ShApp, 0x13709620, 0xC279, 0x11CE, 0xA4,0x9E, 0x44,0x45,0x53,0x54,0x00,0x00)
}

#[allow(non_snake_case)]
mod types {
	use libc::c_void;

	#[repr(C)]
	pub struct GUID {
		pub data1: u32,
		pub data2: u16,
		pub data3: u16,
		pub data4: [u8, ..8],
	}

	pub type BOOL = u32;
	pub type DWORD = u32;
	pub type HRESULT = u32;
	pub type LONG = i32;
	pub type ULONG = u32;
	pub type WORD = u16;

	pub type LCID = DWORD;

	pub type CLSID = GUID;
	pub type FMTID = GUID;
	pub type IID = GUID;

	pub type REFGUID = *const GUID;
	pub type REFCLSID = *const CLSID;
	pub type REFIID = *const IID;
	pub type REFFMTID = *const FMTID;

	pub type DISPID = LONG;
	pub type MEMBERID = DISPID;

	pub type OLECHAR = u16;
	pub type BSTR = *mut OLECHAR;
	pub type LPBSTR = *mut BSTR;

	#[repr(C)]
	pub struct DISPPARAMS {
		rgvarg: *mut VARIANTARG,
		rgdispidNamedArgs: *mut DISPID,
		cArgs: u32,
		cNamedArgs: u32,
	}

	pub struct VARIANT;
	type VARIANTARG = VARIANT;

	pub struct EXCEPINFO;

	pub type ComPtr = *mut c_void;

	#[repr(C)]
	pub struct IUnknown {
		pub __vtable: *mut IUnknown_vtable,
	}

	//pub static IID_IUnknown: &'static str = "00000000-0000-0000-C000-000000000046";
	DEFINE_GUID!(IID_IUnknown, 0x00000000, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)

	#[repr(C)]
	pub struct IUnknown_vtable {
		pub QueryInterface: extern "stdcall" fn(ComPtr, REFIID, *mut ComPtr) -> HRESULT,
		pub AddRef: extern "stdcall" fn(ComPtr) -> ULONG,
		pub Release: extern "stdcall" fn(ComPtr) -> ULONG,
	}

	#[repr(C)]
	pub struct IClassFactory {
		pub __vtable: *mut IClassFactory_vtable,
	}

	//pub static IID_IClassFactory: &'static str = "00000001-0000-0000-C000-000000000046";
	DEFINE_GUID!(IID_IClassFactory, 0x00000001, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46)

	#[repr(C)]
	pub struct IClassFactory_vtable {
		pub __base: IUnknown_vtable,
		pub CreateInstance: extern "stdcall" fn(ComPtr, *mut IUnknown, REFIID, *mut ComPtr) -> HRESULT,
		pub LockServer: extern "stdcall" fn(ComPtr, BOOL) -> HRESULT,
	}

	#[repr(C)]
	pub struct IDispatch {
		pub __vtable: *mut IDispatch_vtable,
	}

	//pub static IID_IDispatch: &'static str = "00020400-0000-0000-C000000000000046";
	DEFINE_GUID!(IID_IDispatch, 0x00020400, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46)

	#[repr(C)]
	pub struct IDispatch_vtable {
		pub __base: IUnknown_vtable,
		pub GetTypeInfoCount: extern "stdcall" fn(ComPtr, *mut u32) -> HRESULT,
		pub GetTypeInfo: extern "stdcall" fn(ComPtr, u32, LCID, *mut *mut ITypeInfo) -> HRESULT,
		pub GetIDsOfNames: extern "stdcall" fn(ComPtr, REFIID, *mut BSTR, u32, LCID, *mut DISPID) -> HRESULT,
		pub Invoke: extern "stdcall" fn(ComPtr, DISPID, REFIID, LCID, WORD, *mut DISPPARAMS, *mut VARIANT, *mut EXCEPINFO, *mut u32) -> HRESULT,
	}

	#[repr(C)]
	pub struct ITypeInfo {
		pub __vtable: *mut ITypeInfo_vtable,
	}

	//pub static IID_ITypeInfo: &'static str = "00020401-0000-0000-C000-000000000046";
	DEFINE_GUID!(IID_ITypeInfo, 0x00020401, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46)

	#[repr(C)]
	//#[idl="oaidl.idl"]
	pub struct ITypeInfo_vtable {
		pub base: IUnknown_vtable,
		// TODO
	}

	#[repr(C)]
	pub struct IShellDispatch {
		pub __vtable: *mut IShellDispatch_vtable,
	}

	//pub static IID_IShellDispatch: &'static str = "d8f015c0-c278-11ce-a49e-444553540000";
	DEFINE_GUID!(IID_IShellDispatch, 0xd8f015c0, 0xc278, 0x11ce, 0xa4,0x9e, 0x44,0x45,0x53,0x54,0x00,0x00)

	#[repr(C)]
	//#[idl="shldisp.idl"]
	pub struct IShellDispatch_vtable {
		pub __base: IDispatch_vtable,
		pub Application: extern "stdcall" fn(ComPtr, *mut *mut IDispatch) -> HRESULT,
		pub Parent: extern "stdcall" fn(ComPtr, *mut *mut IDispatch) -> HRESULT,
		pub NameSpace: *mut c_void, // extern "stdcall" fn(ComPtr, VARIANT, *mut *mut Folder) -> HRESULT,
		pub BrowseForFolder: *mut c_void, // extern "stdcall" fn(ComPtr, u32, BSTR, u32, *mut *mut Folder) -> HRESULT,
		pub Windows: extern "stdcall" fn(ComPtr, *mut *mut IDispatch) -> HRESULT,
		pub Open: *mut c_void, // extern "stdcall" fn(ComPtr, VARIANT) -> HRESULT,
		pub Explore: *mut c_void, // extern "stdcall" fn(ComPtr, VARIANT) -> HRESULT,
		pub MinimizeAll: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub UndoMinimizeALL: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub FileRun: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub CascadeWindows: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub TileVertically: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub TileHorizontally: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub ShutdownWindows: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub Suspend: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub EjectPC: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub SetTime: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub TrayProperties: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub Help: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub FindFiles: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub FindComputer: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub RefreshMenu: extern "stdcall" fn(ComPtr) -> HRESULT,
		pub ControlPanelItem: extern "stdcall" fn(ComPtr, BSTR) -> HRESULT,
	}
}