;cTimer.cls model assembler source

;Runtime patch markers
%define _patch1_    01BCCAABh           ;Relative address to EbMode (vba6.dll or vba5.dll)
%define _patch2_    02BCCAABh           ;Timer ID
%define _patch3_    03BCCAABh           ;Start time
%define _patch4_    04BCCAABh           ;Address of the owner object
%define _patch5_    05BCCAABh           ;Relative address to KillTimer (user32.dll)

;Stack frame
%define dwTime      [esp+20]            ;SystemTime
%define idTimer     [esp+12]            ;Timer ID

[bits 32]
    jmp     _callback                   ;Patched with two nop's if in the IDE
    
            ;Check to see if the IDE is on a breakpoint
    db      0E8h                        ;Far call op-code
    dd      _patch1_                    ;Call EbMode, the relative address to EbMode is patched at runtime
    cmp     eax, 2                      ;If EbMode returns 2 
    je      _return                     ;   The IDE is on a breakpoint
    test    eax, eax                    ;If EbMode returns 0
    je      _kill_tmr                   ;   The IDE has stopped

_callback:  ;Call the owner object's iTimer_Proc interface
    push    _patch2_                    ;Timer ID
    mov     eax, dwTime                 ;Prepare elapsed time calculation
    sub     eax, _patch3_               ;Calculate the elapsed time, patched at runtime
    push    eax                         ;ByVal elapsed time
    mov     eax, _patch4_               ;Address of the owner object, patched at runtime
    push    eax                         ;Push address of the owner object
    mov     eax, [eax]                  ;Get the address of the vTable
    call    dword [eax+1Ch]             ;Call iTimer_Proc, vTable offset 1Ch

_return:    ;Cleanup and exit                                
    ret     16
    
_kill_tmr:  ;The IDE has stopped, kill the timer and return
    mov     ecx, idTimer                ;Get the timer ID
    push    ecx                         ;Push the timer ID
    push    eax                         ;Push 0
    db      0E8h                        ;Far call
    dd      _patch5_                    ;Call KillTimer, patched at runtime
    jmp     _return