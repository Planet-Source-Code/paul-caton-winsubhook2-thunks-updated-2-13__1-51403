;cSubclass.cls model assembler source

;Runtime patch markers
%define _patch1_    01BCCAABh           ;Relative address of the EbMode function
%define _patch2_    02BCCAABh           ;Address of the previous WndProc
%define _patch3_    03BCCAABh           ;Relative address of SetWindowsLong
%define _patch4_    04BCCAABh           ;Table B (before) address
%define _patch5_    05BCCAABh           ;Table B (before) entry count
%define _patch6_    06BCCAABh           ;Address of the previous WndProc
%define _patch7_    07BCCAABh           ;Relative address of CallWindowProc
%define _patch8_    08BCCAABh           ;Table A (after) address
%define _patch9_    09BCCAABh           ;Table A (after) entry count
%define _patchA_    0ABCCAABh           ;Address of the owner object

;Stack frame parameters and local variables. After "push edi" the ebp stack frame will look like this...
%define lParam      [ebp+20]            ;lParam parameter
%define wParam      [ebp+16]            ;wParam parameter
%define uMsg        [ebp+12]            ;Message number parameter
%define hWnd        [ebp +8]            ;Window handle parameter
       ;Information [ebp +4] return address to the caller
       ;Information [ebp +0] previous value of ebp pushed here, (implicitly restored with the leave statement)
%define lReturn     [ebp -4]            ;Local variable lReturn
%define bHandled    [ebp -8]            ;Local variable bHandled
       ;Information [ebp-12] edi saved and restored here
                   
%define GWL_WNDPROC -4                  ;SetWindowsLong WndProc offset

[bits 32]   ;32bit code

;Entry point, setup stack frame
    push    ebp                         ;Preserve base pointer (ebp)
    mov     ebp, esp                    ;Create stack frame
    add     esp, 0FFFFFFF8h             ;Allocate local variable space (2 Longs (8 bytes)) on the stack (esp)
    push    edi                         ;Preserve edi
    
;Initialize locals
    xor     eax, eax                    ;Clear eax
    mov     lReturn, eax                ;Clear lReturn
    mov     bHandled, eax               ;Clear bHandled

    jmp     _no_ide                     ;Patched with two nop's if running in the the IDE

;Check to see if the IDE is on a breakpoint or has stopped
    db      0E8h                        ;Far call
    dd      _patch1_                    ;Call EbMode, patched at runtime
    cmp     eax, 2                      ;If 2 
    je      _break                      ;  IDE is on a breakpoint, just call the original WndProc
    test    eax, eax                    ;If 0
    je      _unsub                      ;  IDE has stopped, unsubclass the window
    
_no_ide:
    call    _before                     ;Before processing
    cmp     bHandled, dword 0           ;Has message been handled?
    jne     _return                     ;  Yep, return
    call    _original                   ;Call the original 
    call    _after                      ;After processing

_return:    ;Cleanup and exit
    pop     edi                         ;Restore edi
    mov     eax, lReturn                ;Function return value
    leave                               ;Restore Stack for Procedure Exit
    ret     16                          ;Return and adjust esp

_break:     ;The IDE is on a breakpoint, call the original WndProc and return
    call    _original
    jmp     _return

_unsub:     ;IDE has stopped, unsubclass the window
    push    _patch2_                    ;Address of the previous WndProc, patched at runtime
    push    dword GWL_WNDPROC           ;WndProc index
    push    dword hWnd                  ;Push the window handle
    db      0e8h                        ;Far call
    dd      _patch3_                    ;Relative address of SetWindowsLong, patched at runtime
    jmp     _return                     ;Return
    
_before:    ;Callback before the original WndProc
    xor     edx, edx                    ;edx = 0
    dec     edx                         ;edx = -1, bBefore = True
    mov     edi, _patch4_               ;Table B (before) address, patched at runtime
    mov     ecx, _patch5_               ;Table B (before) entry count, patched at runtime
    call    _callback                   ;Callback before
    ret
    
_original:  ;Call original WndProc
    push    dword lParam                ;ByVal lParam
    push    dword wParam                ;ByVal wParam
    push    dword uMsg                  ;ByVal uMsg
    push    dword hWnd                  ;ByVal hWnd
    push    _patch6_                    ;Address of the previous WndProc, patched at runtime
    db      0e8h                        ;Far Call 
    dd      _patch7_                    ;Relative address of CallWindowProc, patched at runtime
    mov     lReturn, eax                ;Preserve the return value
    ret
    
_after:     ;Callback after the original WndProc
    xor     edx, edx                    ;edx = 0, bBefore = False (After)
    mov     edi, _patch8_               ;Table A (after) address, patched at runtime
    mov     ecx, _patch9_               ;Table A (after) entry count, patched at runtime
    call    _callback                   ;Callback after
    ret
    
_callback   ;Callback, edx indicates before or after (-1 or 0, True or False)
    jecxz   _skip                       ;Entry count (ecx) = 0, just skip
    or      ecx, ecx                    ;Set flags
    js      _call                       ;Entry count is negative, all messages callback to iSubClass_Proc
    
;Scan the table for a matching entry    
    mov     eax, uMsg                   ;Message number to search for
    repne   scasd                       ;Scan the table
    jne     _skip                       ;If the uMsg number isn't found in the table just skip

_call:      ;Callback to the owners iSubclass_Proc implemented interface
    lea     eax, lParam                 ;Address of lParam into eax
    push    eax                         ;Push ByRef lParam
    lea     eax, wParam                 ;Address of wParam into eax
    push    eax                         ;Push ByRef wParam
    lea     eax, uMsg                   ;Address of uMsg into eax
    push    eax                         ;Push ByRef uMsg
    lea     eax, hWnd                   ;Address of hWnd into eax
    push    eax                         ;Push ByRef hWnd
    lea     eax, lReturn                ;Address of lReturn into eax
    push    eax                         ;Push ByRef lReturn
    lea     eax, bHandled               ;Address of bHandled into eax
    push    eax                         ;Push ByRef bHandled
    push    edx                         ;Push ByVal bBefore

    mov     eax, _patchA_               ;Address of the owner object, patched at runtime
    push    eax                         ;Push address of the owner object
    mov     eax, [eax]                  ;Get the address of the vTable into eax
    call    dword [eax+1Ch]             ;Call iSubclass_Proc, vTable offset 1Ch
_skip:    
    ret
    