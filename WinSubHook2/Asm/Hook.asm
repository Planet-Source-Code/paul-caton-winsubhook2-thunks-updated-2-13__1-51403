;cHook.cls model assembler source

;Runtime patch markers
%define _patch1_    01BCCAABh           ;Relative address of the EbMode function
%define _patch2_    02BCCAABh           ;Hook handle for UnhookWindowsHookEx
%define _patch3_    03BCCAABh           ;Relative address of UnhookWindowsHookEx
%define _patch4_    04BCCAABh           ;Hook handle for CallNextHookEx
%define _patch5_    05BCCAABh           ;Relative address of CallNextHookEx
%define _patch6_    06BCCAABh           ;Address of the owner object

;Stack frame parameters and local variables
%define lParam      [ebp+16]            ;lParam
%define wParam      [ebp+12]            ;wParam
%define nCode       [ebp +8]            ;Hook code
%define lReturn     [ebp -4]            ;Local lReturn
%define bHandled    [ebp -8]            ;Local bHandled

[bits 32]   ;32 bit code
    push    ebp                         ;Preserve ebp
    mov     ebp, esp                    ;Create stack frame
    add     esp, 0FFFFFFF8h             ;Allocate local variable space on the stack

;Initialize locals
    xor     edx, edx                    ;Clear edx
    mov     lReturn, edx                ;Clear lReturn
    mov     bHandled, edx               ;Clear bHandled

    jmp     _no_ide                     ;Patched with two nop's if in the IDE

;Check to see if the IDE is on a breakpoint or has stopped
    db      0E8h                        ;Far call
    dd      _patch1_                    ;Call EbMode, patched at runtime
    cmp     eax, 2                      ;If 2 
    je      _break                      ;  IDE is on a breakpoint, call the next hook
    test    eax, eax                    ;If 0
    je      _unhook                     ;  IDE has stopped, unsubclass the window

_no_ide:    ;Preserve registers
    call    _before                     ;Before processing
    cmp     bHandled, dword 0           ;Has message been handled?
    jne     _return                     ;  Yep, return
    call    _next_hook                  ;Call the original 
    call    _after                      ;After processing

_return:    ;Cleanup and exit
    mov     eax, lReturn                ;Function return value
    leave                               ;Restore Stack for Procedure Exit
    ret     12                          ;Return and adjust esp

_break:     ;We're on a breakpoint, call the next hook and return
    call    _next_hook
    jmp     _return

_unhook:    ;The IDE has stopped, unhook and return
    push    _patch2_                    ;Current hook handle, patched at runtime
    db      0e8h                        ;Far call
    dd      _patch3_                    ;Relative call address of UnhookWindowsHookEx, patched at runtime
    jmp     _return                     ;Return
    
_before:    ;Callback before        
    xor     edx, edx
    dec     edx                         ;edx = -1, bBefore = True
    call    _callback                   ;Callback before
    ret
    
_next_hook: ;Call next hook    
    push    dword lParam                ;ByVal lParam
    push    dword wParam                ;ByVal wParam
    push    dword nCode                 ;ByVal nCode
    push    _patch4_                    ;ByVal hook handle, patched at runtime
    db      0e8h                        ;Far call
    dd      _patch5_                    ;Relative address of CallNextHookEx, patched at runtime
    mov     lReturn, eax                ;Preserve the CallNextHookEx return value
    ret
    
_after:     ;Callback after
    xor     edx, edx                    ;edx = 0, bBefore = False (After)
    call    _callback                   ;Callback after
    ret
    
_callback   ;Callback to the owners iHook_Proc implemented interface
    lea     eax, lParam                 ;Address of lParam 
    push    eax                         ;push ByRef lParam
    lea     eax, wParam                 ;Address of wParam
    push    eax                         ;push ByRef wParam
    lea     eax, nCode                  ;Address of nCode
    push    eax                         ;push ByRef nCode
    lea     eax, lReturn                ;Address of lReturn
    push    eax                         ;push ByRef lReturn
    lea     eax, bHandled               ;Address of bHandled 
    push    eax                         ;push ByRef bHandled
    push    edx                         ;push ByVal bBefore

    mov     eax, _patch6_               ;Address of the owner object, patched at runtime
    push    eax                         ;Push address of the owner object
    mov     eax, [eax]                  ;Get the address of the vTable
    call    dword [eax+1Ch]             ;Call iSubclass_Proc, vTable offset 1Ch
_skip:    
    ret
