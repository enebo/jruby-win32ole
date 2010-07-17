require 'ffi'

module Win
  module User32
    extend FFI::Library
    ffi_lib 'user32'
    ffi_convention :stdcall

    PM_REMOVE = 1

    DWORD = :uint32
    HWND = :pointer
    UINT = :uint32
    LPARAM = :pointer # 32 or 64 bit int actually
    WPARAM = :pointer # 32 or 64 bit int actually

    class POINT < FFI::Struct
      layout :x, :long, :y, :long
    end

    class MSG < FFI::Struct
      layout :hwnd, HWND,
        :message, UINT,
        :wParam, WPARAM,
        :lParam, LPARAM,
        :time, DWORD,
        :pt, POINT  
    end

    # BOOL WINAPI PeekMessage(out LPMSG lpMsg, in_opt HWND hWnd, 
    #   in UINT wMsgFilterMin, in UINT wMsgFilterMax, in UINT wRemoveMsg)
    attach_function :peek_message, :PeekMessageA, [:pointer, HWND, UINT, UINT, UINT], :bool

    # BOOL WINAPI TranslateMessage(in const MSG *lpMsg)
    attach_function :translate_message, :TranslateMessage, [:pointer], :bool

    # LRESULT WINAPI DispatchMessage(__in  const MSG *lpmsg)
    attach_function :dispatch_message, :DispatchMessageA, [:pointer], :pointer
  end
end
