/****************************************************************
 *   Prog: The Last Crusade										*
 *     By: Peter S. VanLund										*
 *   Desc: All code written by Peter S. VanLund.				*
 *         Gameplay design, story, & sounds by Patrick Dwyer.	*
 ****************************************************************/

#include "Game.h"
#include <windows.h>

Game game; // Game engine object

LRESULT CALLBACK WndProc(HWND,UINT,WPARAM,LPARAM); // Windows callback

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
				   PSTR szCmdLine, int iCmdShow)
{
	// If a map file is passed in on commandline, load as custom map
	if(strlen(szCmdLine)>0)
	{
		game.addMap(szCmdLine);
		game.useCustom();
	}
	// Otherwise, play game normally
	else
	{
		game.addMap("forest.map");
		game.addMap("graveyard.map");
		game.addMap("castle.map");
	}
	// Windows jargon
	char* szAppName = "LastCrusade";
	HWND hwnd;
	MSG msg;
	WNDCLASS wndclass;

	wndclass.style = CS_HREDRAW | CS_VREDRAW;
	wndclass.lpfnWndProc = WndProc;
	wndclass.cbClsExtra = 0;
	wndclass.cbWndExtra = 0;
	wndclass.hInstance = hInstance;
	wndclass.hIcon = LoadIcon(NULL,IDI_APPLICATION);
	wndclass.hCursor = LoadCursor(NULL,IDC_ARROW);
	wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wndclass.lpszMenuName = NULL;
	wndclass.lpszClassName = szAppName;

	if(!RegisterClass(&wndclass))
	{
		MessageBox(NULL,"This program requires Windows NT!",szAppName,MB_ICONERROR);
		return 0;
	}

	hwnd = CreateWindow(szAppName,				// window class name
		                "The Last Crusade",		// window caption
						WS_OVERLAPPEDWINDOW,	// window style
						CW_USEDEFAULT,			// initial x position
						CW_USEDEFAULT,			// initial y position
						400,					// initial x size
						200,					// initial y size
						NULL,					// parent window handle
						NULL,					// window menu handle
						hInstance,				// program instance handle
						NULL);					// creation parameters
	ShowWindow(hwnd,iCmdShow);
	UpdateWindow(hwnd);

	// Start game
	game.play();

	while(GetMessage(&msg,NULL,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return msg.wParam;
}

// Handles Windows messages
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	HDC hdc;
	PAINTSTRUCT ps;
	RECT rect;

	switch(message)
	{
		case WM_CREATE:
			return 0;
		case WM_KEYUP:
			// Send key to engine
			game.processKey(wParam);
			return 0;
		case WM_PAINT:
			hdc = BeginPaint(hwnd,&ps);
			GetClientRect(hwnd,&rect);
			DrawText(hdc,"The Last Crusade",-1,&rect,DT_SINGLELINE | DT_CENTER | DT_VCENTER);
			EndPaint(hwnd,&ps);
			return 0;
		case WM_DESTROY:
			PostQuitMessage(0);
			exit(0);
			return 0;
	}
	return DefWindowProc(hwnd,message,wParam,lParam);
}