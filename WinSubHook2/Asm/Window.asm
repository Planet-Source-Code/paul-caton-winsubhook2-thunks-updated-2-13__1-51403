;cWindow.cls model assembler source

;Runtime patch markers
%define _patch1_ 01BCCAABh              ;Relative address of the EbMode function
%define _patch2_ 02BCCAABh              ;Table B (before) entry count
%define _patch3_ 03BCCAABh              ;Table B (before) address
%define _patch4_ 04BCCAABh              ;Object address of the owner object
%define _patch5_ 05BCCAABh              ;Relative address of DefWindowProc
%define _patch6_ 06BCCAABh              ;Relative address of DestroyWindow
%define _patch7_ 07BCCAABh              ;Application instance handle
%define _patch8_ 08BCCAABh              ;Address of the class name
%define _patch9_ 09BCCAABh              ;Relative address of UnregisterClass function
%define _patchA_ 0ABCCAABh              ;Marks the location of the class name

;Stack frame parameters and local variables
%define lParam   [ebp+20]               ; lParam
%define wParam   [ebp+16]               ; wParam
%define uMsg     [ebp+12]               ; Message number
%define hWnd     [ebp+ 8]               ; Window handle
%define lReturn  [ebp- 4]               ; Local lReturn
%define bHandled [ebp- 8]               ; Local bHandled

[bits 32]   ;Entry point, setup stack frame
    push    ebp                         ;Preserve ebp
    mov     ebp, esp                    ;Create stack frame
    add     esp, 0FFFFFFF8h             ;Allocate local variable space on the stack
    push    edi                         ;Preserve edi

;Initialize locals
    xor     eax, eax                    ;Clear eax
    mov     lReturn, eax                ;Clear lReturn
    mov     bHandled, eax               ;Clear bHandled

    jmp     _no_ide                     ;Patched with two nop's if in the IDE

;Check to see if the IDE is on a breakpoint or has stopped
    db      0E8h                        ;Far call
    dd      _patch1_                    ;Call EbMode, patched at runtime
    cmp     eax, 2                      ;If 2 
    je      _prev_proc                  ;  IDE is on a breakpoint, just call the previous WndProc
    test    eax, eax                    ;If 0
    je      _destroy                    ;  IDE has stopped, destroy the window
    
_no_ide:    ;Check the table entry count
    mov     ecx, _patch2_               ;Table entry count, patched at runtime
    jecxz   _prev_proc                  ;Entry count = 0, call default processing
    or      ecx, ecx                    ;Check for a negative entry count
    js      _callback                   ;Negative, all messages          

;Scan the table for a matching entry    
    mov     edi, _patch3_               ;Table address, patched at runtime
    mov     eax, uMsg                   ;Message number to search for
    repne   scasd                       ;Scan the table
    jne     _prev_proc                  ;Call default processing

_callback:  ;Callback to the owners iWindow_Proc implemented interface
    lea     eax, lParam                 ;Address of lParam 
    push    eax                         ;Push ByRef lParam
    lea     eax, wParam                 ;Address of wParam
    push    eax                         ;Push ByRef wParam
    lea     eax, uMsg                   ;Address of uMsg 
    push    eax                         ;Push ByRef uMsg
    lea     eax, hWnd                   ;Address of hWnd 
    push    eax                         ;Push ByRef hWnd
    lea     eax, lReturn                ;Address of lReturn 
    push    eax                         ;Push ByRef lReturn
    lea     eax, bHandled               ;Address of bHandled 
    push    eax                         ;Push ByRef bHandled

    mov     eax, _patch4_               ;Address of the owner object, patched at runtime
    push    eax                         ;Push address of the owner object
    mov     eax, [eax]                  ;Get the address of the vTable
    call    dword [eax+1Ch]             ;Call iWindow_Proc, vTable offset 1Ch
    cmp     bHandled, dword 0           ;Has message been handled?
    jne     _return                     ;Yep, return
    
_prev_proc: ;Call the previous window proc
    push    dword lParam                ;ByVal lParam
    push    dword wParam                ;ByVal wParam
    push    dword uMsg                  ;ByVal uMsg
    push    dword hWnd                  ;ByVal hWin
    db      0e8h                        ;Far call
    dd      _patch5_                    ;Relative address of DefWindowProc, patched at runtime
    mov     lReturn, eax                ;Preserve the return value

_return:    ;Cleanup and exit
    pop     edi                         ;Restore edi
    mov     eax, lReturn                ;Function return value
    leave                               ;Restore Stack for Procedure Exit
    ret     16                          ;Return and adjust esp
    
_destroy:   ;vtable is gone, destroy the window, attempt to unregister the window class (will only work on the last window) and depart
    push    dword hWnd                  ;Push the window handle
    db      0e8h                        ;Far call
    dd      _patch6_                    ;Relative address of DestroyWindow, patched at runtime

    push    _patch7_                    ;Application hInstance
    push    _patch8_                    ;Adress of the class name
    db      0e8h                        ;Far call
    dd      _patch9_                    ;Relative address of UnregisterClass, patched at runtime
    jmp     _return                     ;Return
    _Class  dd  _patchA_                ;Marker for the location of the class name