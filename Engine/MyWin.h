/****************************************************************
 *  Class: MyWin												*
 *     By: Peter S. VanLund										*
 *   Desc: Special routines using Windows, mainly for input.	*
 ****************************************************************/

class MyWin
{
	public:
		static void DoEvents();		// Allow Windows messages to be handled
		static WPARAM getKey();		// Get keystroke
		static WPARAM getYorN();	// Get 'Y' or 'N' keystrokes
		static WPARAM getYorNRun();	// Get 'Y'('R') or 'N'('A') keystrokes
		static WPARAM get123();		// Get '1', '2', or '3' keystrokes
};

// Allow Windows messages to be handled
void MyWin::DoEvents()
{
	MSG msg;
	while(PeekMessage(&msg,NULL,0,0,PM_REMOVE))
	{
		TranslateMessage(&msg);
		// Do not dispatch keystrokes
		if(msg.message!=WM_KEYUP) DispatchMessage(&msg);
	}
}

// Get keystroke
WPARAM MyWin::getKey()
{
	MSG msg;
	while(PeekMessage(&msg,NULL,0,0,PM_REMOVE))
	{
		TranslateMessage(&msg);
		if(msg.message==WM_KEYUP)
			return msg.wParam;
		DispatchMessage(&msg);
	}
	return 0;
}

// Get 'Y' or 'N' keystrokes
WPARAM MyWin::getYorN()
{
	WPARAM key = 0;
	while(key!='Y' && key!='N')
		key = getKey();
	return key;
}

// Get 'Y'('R') or 'N'('A') keystrokes
WPARAM MyWin::getYorNRun()
{
	WPARAM key = 0;
	while(key!='Y' && key!='N')
	{
		key = getKey();
		key = (key=='R') ? 'Y' : (key=='A') ? 'N' : key;
	}
	return key;
}

// Get '1', '2', or '3' keystrokes
WPARAM MyWin::get123()
{
	WPARAM key = 0;
	while(key!='1' && key!='2' && key!='3')
	{
		key = getKey();
	}
	return key;
}