/********************************************************
 *  Class: Sound										*
 *     By: Peter S. VanLund								*
 *   Desc: Contains all methods for creating & playing	*
 *         sounds (spatial or traditional stereo) with	*
 *         FMOD.										*
 ********************************************************/


#include <fmod.h>
#include <fmod_errors.h>
#include <windows.h>
#include <vector>
#include "MyWin.h"

using namespace std;

class Sound
{
	private:
		FSOUND_SAMPLE *sound;				// FMOD sample
		int channel;						// FMOD channel
		int posX, posY;						// Position of sound
		bool spatial;						// Spatial boolen
		static int listenerX, listenerY;	// Listener location
		string filepath;					// Path to sound file
	public:
		Sound();							// Blank constructor
		Sound(string);						// Constructor taking path
		Sound(string,int,int);				// Constructor taking path & pos
		~Sound();							// Destructor
		static void init();					// FMOD initialization
		static void shutdown();				// FMOD shutdown
		static void setListener(int,int);	// Set listener position
		bool operator==(Sound);				// Equality of sounds
		FSOUND_SAMPLE* getSample();			// Get FMOD sample
		int getChannel();					// Get FMOD channel
		int getPosX();						// Get x-coord position
		int getPosY();						// Get y-coord position
		string getPath();					// Get path to sound file
		bool isSpatial();					// Spatial test
		void play(bool);					// Play sound
		bool playAndWait(bool);				// Play and wait until finish
		WPARAM playAndGetYorN();			// Play sound and get input
		WPARAM playAndGetYorNRun();			// Play sound and get input
		WPARAM playAndGet123();				// Play sound and get input
		void stop();						// Stop sound
		void fadeOut();						// Fade sound out
		bool isPlaying();					// Playing test
		void setVolume(int);				// Set sound volume
};

// Set initial listener position
int Sound::listenerX = 0;
int Sound::listenerY = 0;

// Blank constructor
Sound::Sound(){}

//////////////////////////////////////////////////////////////////////////////////////////
//         Constructor given path: Loads sample for traditional stereo playing          //
//////////////////////////////////////////////////////////////////////////////////////////
Sound::Sound(string path)																//
{																						//
	filepath = path;																	//
	channel = -1;																		//
	// Load sample																		//
	sound = FSOUND_Sample_Load(FSOUND_FREE,path.c_str(),FSOUND_HW2D,0,0);				//
																						//
	// If problem, report error															//
	if(!sound)																			//
	{																					//
		// DEBUG:																		//
		//MessageBox(NULL,FMOD_ErrorString(FSOUND_GetError()), "Sound File Error", 0);	//
		return;																			//
	}																					//
	// Set sound properties																//
	spatial = false;																	//
}																						//
//////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////
//         Constructor given path & position: Loads sample for spatial playing          //
//////////////////////////////////////////////////////////////////////////////////////////
Sound::Sound(string path, int x, int y)													//
{																						//
	filepath = path;																	//
	channel = -1;																		//
	// Load sample																		//
	sound = FSOUND_Sample_Load(FSOUND_FREE,path.c_str(),FSOUND_HW3D,0,0);				//
																						//
	// If problem, report error															//
	if(!sound)																			//
	{																					//
		// DEBUG:																		//
		//MessageBox(NULL,FMOD_ErrorString(FSOUND_GetError()), "Sound File Error", 0);	//
		return;																			//
	}																					//
	// Set sound properties																//
	FSOUND_Sample_SetMinMaxDistance(sound,0.0f,10000.0f);								//
	spatial = true;																		//
	posX = x; posY = y;																	//
}																						//
//////////////////////////////////////////////////////////////////////////////////////////

Sound::~Sound()
{
	if(sound) FSOUND_Sample_Free(sound);
};

//////////////////////////////////////////////////
//           init & shutdown of FMOD            //
//////////////////////////////////////////////////
void Sound::init(){FSOUND_Init(44100,16,0);}	//
void Sound::shutdown(){FSOUND_Close();}			//
//////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////
//                     setListener: Sets listener position                      //
//////////////////////////////////////////////////////////////////////////////////
void Sound::setListener(int x, int y)											//
{																				//
	listenerX=x;listenerY=y;													//
	float pos[] = {listenerX,0,listenerY};										//
	float vel[] = {0,0,0};														//
	FSOUND_3D_Listener_SetAttributes(pos,vel,0.0f,0.0f,1.0f,0.0f,1.0f,0.0f);	//
	FSOUND_Update();															//
}																				//
//////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////
//                     operator==: Tests equality of sounds                     //
//////////////////////////////////////////////////////////////////////////////////
bool Sound::operator==(Sound s)													//
{																				//
	return (sound==s.getSample() && channel==s.getChannel() &&					//
		    posX==s.getPosX() && posY==s.getPosY() && spatial==s.isSpatial());	//
}																				//
//////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////
//             Sound querying functions             //
//////////////////////////////////////////////////////
FSOUND_SAMPLE* Sound::getSample(){return sound;}	//
int Sound::getChannel(){return channel;}			//
int Sound::getPosX(){return posX;}					//
int Sound::getPosY(){return posY;}					//
string Sound::getPath(){return filepath;}			//
bool Sound::isSpatial(){return spatial;}			//
//////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
//                  play: Plays sound & may be looped                   //
//////////////////////////////////////////////////////////////////////////
void Sound::play(bool loop)												//
{																		//
	if(sound && !isPlaying())											//
	{																	//
		if(loop) FSOUND_Sample_SetMode(sound,FSOUND_LOOP_NORMAL);		//
		else     FSOUND_Sample_SetMode(sound,FSOUND_LOOP_OFF);			//
		stop();															//
		channel = FSOUND_PlaySoundEx(FSOUND_FREE,sound,NULL,spatial);	//
		if(!spatial) FSOUND_SetVolume(channel, 255);					//
		else															//
		{																//
			FSOUND_SetVolume(channel,32);								//
			float pos[] = {posX,0,posY};								//
			float vel[] = {0,0,0};										//
			FSOUND_3D_SetAttributes(channel,pos,vel);					//
			FSOUND_SetPaused(channel,false);							//
		}																//
	}																	//
}																		//
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////
// playAndWait: Play & wait for finish	//
//////////////////////////////////////////
bool Sound::playAndWait(bool skip)		//
{										//
	// Return whether sound was skipped	//
	bool ret = false;					//
	play(false);						//
	while(isPlaying())					//
	{									//
		WPARAM key = MyWin::getKey();	//
		if(skip && key=='F')			//
		{								//
			ret = true;					//
			break;						//
		}								//
	}									//
	if(isPlaying()) stop();				//
	return ret;							//
}										//
//////////////////////////////////////////

//////////////////////////////////////////
// playAndGet[KEY]: Play and get [KEY]  //
//////////////////////////////////////////
WPARAM Sound::playAndGetYorN()			//
{										//
	play(false);						//
	WPARAM key = MyWin::getYorN();		//
	if(isPlaying()) stop();				//
	return key;							//
}										//
//////////////////////////////////////////
WPARAM Sound::playAndGetYorNRun()		//
{										//
	play(false);						//
	WPARAM key = MyWin::getYorNRun();	//
	if(isPlaying()) stop();				//
	return key;							//
}										//
//////////////////////////////////////////
WPARAM Sound::playAndGet123()			//
{										//
	play(false);						//
	WPARAM key = MyWin::get123();		//
	if(isPlaying()) stop();				//
	return key;							//
}										//
//////////////////////////////////////////

//////////////////////////////////
//      stop: Stops sound       //
//////////////////////////////////
void Sound::stop()				//
{								//
	if(channel==-1) return;		//
	FSOUND_StopSound(channel);	//
	channel = -1;				//
}								//
//////////////////////////////////

//////////////////////////////////////////////
//        fadeOut: Fades a sound out        //
//////////////////////////////////////////////
void Sound::fadeOut()						//
{											//
	if(channel==-1) return;					//
	int vol = FSOUND_GetVolume(channel);	//
	int dec = vol/10;						//
	while(vol>0)							//
	{										//
		vol-=dec;							//
		setVolume(vol);						//
		Sleep(150);							//
		MyWin::DoEvents();					//
	}										//
	if(isPlaying()) stop();					//
}											//
//////////////////////////////////////////////

//////////////////////////////////////////
//     isPlaying: Is sound playing?     //
//////////////////////////////////////////
bool Sound::isPlaying()					//
{										//
	if(channel==-1) return false;		//
	return FSOUND_IsPlaying(channel)>0;	//
}										//
//////////////////////////////////////////

//////////////////////////////////////////////////////
//           setVolume: Set sound volume            //
//////////////////////////////////////////////////////
void Sound::setVolume(int v)						//
{													//
	if(channel!=-1) FSOUND_SetVolume(channel, v);	//
}													//
//////////////////////////////////////////////////////
