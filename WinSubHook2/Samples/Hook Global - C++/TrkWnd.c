//
// This now just an "educational" sample. HooKey uses a previously undocumented api call to achieve the same end.
// Just a couple of usage notes, seeing as you now can't see how HooKey uses this DLL.
//
// In your VB app...
// Create a unique message number using the RegisterWindowMessage api
// Setup a cSubclasser instance
// AddMsg the unique message number, MSG_BEFORE
// Call HookStart
// Do whatever in the implemented interface callback
// AND don't forget to call HookStop when you're through - Though I've added code to uninstall the hook
// should the app that called HookStart stop... let's start as we mean to carry on.. with care.
//

//
// TrkWnd.dll installs a system-wide CBT hook into all running window processes. When the hook receives an app activation
// notification it posts a message to the application that installed the hook. The message includes
// the Window handle of the application that's just about to activate plus whether the mouse was used to effect the change.
//
// Why? HooKey does a good job with the system-wide keyboard hook, it's no huge deal to capture global keystrokes in a VB 
// application. However, it's all very well to know what's being typed, but what app is it being typed into? For that we need 
// a hook that resides in a *true* dll... not the com oddity that VB produces. So for now we need the assistance of Visual C.
//
// With regard to HooKey and Planet-Source-Code's restriction on binary uploads... and bearing in mind that some of you
// wont have access to Visual C++ 6 - I will package the dll into HooKey's resource file and arrange for HooKey to write
// TrkWnd.dll to the application directory if it isn't found.
//
// Paul_Caton@hotmail.com
// Copyright free, use and abuse as you see fit.
//

#define STRICT
#define WINDOWS_LEAN_AND_MEAN
#include <windows.h>

//Ordinarily, module level variables are unique to the instance of the dll. Here we create a 
//shared memory segment such that every instance of the dll will be sharing the same values.
#pragma data_seg( ".shared" )
	volatile HINSTANCE hLoadInst	= 0;	//Instance handle of the app (HooKey) that initially loaded us.
	volatile HHOOK     hHookPrev	= 0;	//Handle of the previous hook.
	volatile HWND	   hWndNotify	= 0;	//The handle of the window to receive notification.
	volatile UINT      uMsgNotify	= 0;	//The message to send to the the window receiving notifications.
#pragma data_seg()
#pragma comment(linker, "/SECTION:.shared,RWS")

//CBT Windows hook callback function
//Note: After the hook is set, Windows injects this dll into the execution space of every process with a Windows interface,
//already running or yet to be started. Because we are running everywhere.. it's important to get our work done as quickly 
//as possible so as not to observably affect the performance of the system.
LRESULT CALLBACK CbtProc(int    nCode,		//Hook code
						 WPARAM wParam,		//Hook dependant value. With CBT/HCBT_ACTIVATE wParam is the Windows handle of the app that is being activated
						 LPARAM lParam)		//Hook dependant value. With CBT/HCBT_ACTIVATE lParam is a pointer to a CBTACTIVATESTRUCT
{
	//If an application is activating
	if (nCode == HCBT_ACTIVATE)
	{
		//Post a message to the app that HookStart'ed us. wParam is the 
		//window handle of the app that is activating. lParam points to
		//CBTACTIVATESTRUCT struct, which apart from the hWnd (which we already
		//know) indicates whether the mouse was used to active the app. 
		//Handily, the fMouse Boolean is the first member of the struct and thus
		//is pointed to by lParam. We use PostMessage instead of SendMessage
		//because Post returns immediately after adding the message to the
		//appropriate event queue whereas Send wouldn't return until the message
		//wass digested by HooKey.
		PostMessage(hWndNotify, uMsgNotify, wParam, *((LPARAM *)lParam));
	}
	
	//Call the next hook in the chain
	return CallNextHookEx(hHookPrev, nCode, wParam, lParam);
}

//Start the CBT hook, called externally by HooKey
__declspec(dllexport) int __stdcall
HookStart(HWND hWnd,	//This is the window handle to send notifications to
		  UINT uMsg)	//This is the message number to send
{  
	//If the hook is already installed.
	if (hHookPrev)   
		return -1;  

	//Save these values into our dll shared memory
	hWndNotify = hWnd;	//Windows handle of the application to notify of application activations
	uMsgNotify = uMsg;  //The message number to send to the notification window

	//Create the CBT hook
	hHookPrev = SetWindowsHookEx(WH_CBT, CbtProc, hLoadInst, 0);  
	
	if (!hHookPrev)    
		return -1;  //Failed to set the hook
	
	return 0;		//Success
}

//Stop the CBT hook, called externally by HooKey
__declspec(dllexport) int __stdcall
HookStop()
{  
	//If hook is set
	if (hHookPrev != 0)
	{
		//Unhook
		if (UnhookWindowsHookEx(hHookPrev))
		{
			hHookPrev = 0;	//Indicate that we're unhooked
			return 0;		//Success
		}
	}

	return -1;
}

//Low level dll startup - we're not using any C runtime library functions
//so by providing out own _DllMainCRTStartup we avoid the C runtime initializaion
//that ensues if the linker links in the default _DllMainCRTStartup
BOOL WINAPI _DllMainCRTStartup(HINSTANCE hInst, DWORD reason, LPVOID reserved)
{  
	if (reason == DLL_PROCESS_ATTACH)  
	{    
		//We don't need no stinkin' thread library calls
		DisableThreadLibraryCalls(hInst); 

		//If the shared memory variable hLoadInst = 0 then the app that is loading
		//us is the one we'll be reporting to. So... supposing that app unloaded
		//neglecting to call HookStop - we can clean up for him. See DLL_PROCESS_DETACH
		if (hLoadInst == 0) 
			hLoadInst = hInst;

	} else if (reason == DLL_PROCESS_DETACH)
	{
		//If the app that started us is detaching from this instance of the dll
		if (hInst == hLoadInst)
		{
			//If the hook still exists
			if (hHookPrev != 0) 
				HookStop();	//Stop it!
		}
	}

	return TRUE;
}