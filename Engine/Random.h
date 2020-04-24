/********************************************************
 *  Class: Random										*
 *     By: Peter S. VanLund								*
 *   Desc: Contains random methods for game				*
 ********************************************************/

#include <time.h>

class Random
{
	public:
		static void  init();				// Initialize random number generator
		static short getShort(short,short);	// Get short between min & max (inclusive)
		static bool  checkOdds(short);		// Return true odds % of the time
};

//////////////////////////
// init: Initalize seed //
//////////////////////////
void Random::init()		//
{						//
	srand(time(NULL));	//
}						//
//////////////////////////

//////////////////////////////////////////////////
//  getShort: Get short in [min,max] inclusive  //
//////////////////////////////////////////////////
short Random::getShort(short min, short max)	//
{												//
	return (rand()%(max-min+1))+min;			//
}												//
//////////////////////////////////////////////////

//////////////////////////////////////
// checkOdds: True odds% of the time//
//////////////////////////////////////
bool Random::checkOdds(short odds)	//
{									//
	return (rand()%100)<odds;		//
}									//
//////////////////////////////////////