//////////////////////////////////////////////////////////////////////////
//                      Map File Data Structures						//
//////////////////////////////////////////////////////////////////////////
struct MAPHEADER														//
{																		//
	short mStart;			// Start node								//
	short mEnd;				// End node									//
	short mNodeCount;		// # of nodes								//
	short mNPCCount;		// # of NPCs								//
	short mSoundCount;		// # of sounds (events | dialogue | music)	//
	short mItemCount;		// # of items								//
};																		//
																		//
struct NODEHEADER														//
{																		//
	short nIndex;			// Index in VB								//
	short nNorth;			// Node to north							//
	short nSouth;			// Node to south							//
	short nEast;			// Node to east								//
	short nWest;			// Node to west								//
	short nMusicCount;		// # of songs for node						//
	short nItemCount;		// # of items for node						//
	short nNPCCount;		// # of NPCs for node						//
	short nReqItemCount;	// # of required items						//
	char  nImage[20];		// Image for node							//
};																		//
																		//
struct NODESOUND														//
{																		//
	short nSound;			// Sound #									//
};																		//
																		//
struct NODEITEM															//
{																		//
	short nItem;			// Item #									//
	short nPercent;			// Chance itel will be at node				//
};																		//
																		//
struct NODENPC															//
{																		//
	short nNPC;				// NPC #									//
	short nPercent;			// Chance NPC will be at node				//
};																		//
																		//
struct NODEREQITEM														//
{																		//
	short nReqItem;			// Item #									//
};																		//
																		//
struct SOUNDDATA														//
{																		//
	char  sName[20];		// MP3 filename								//
	bool  sSpatial;			// Boolean for spatial						//
	short sNode;			// Node source								//
	short sXCoord;			// Spatial x coord							//
	short sYCoord;			// Spatial y coord							//
};																		//
																		//
struct ITEMDATA															//
{																		//
	char  iName[20];		// Item name								//
	short iType;			// Weapon | Armor | Potion | Gold | Special	//
	short iValue;			// Item value								//
	short iNameSound;		// Name sound								//
	short iActionSound;		// Action sound								//
};																		//
																		//
struct NPCDATA															//
{																		//
	char  cName[20];		// NPC name									//
	short cType;			// Enemy | Friend | Vendor | Leprechaun		//
	short cStrMin;			// Strength min								//
	short cStrMax;			// Strength max								//
	short cDefMin;			// Defense min								//
	short cDefMax;			// Defense max								//
	short cHPMin;			// HP min									//
	short cHPMax;			// HP max									//
	short cRunPerc;			// Run %									//
	short cNameSound;		// Name sound								//
	short cActionSound;		// Action sound								//
	short cItemCount;		// # of items								//
};																		//
																		//
struct NPCITEM															//
{																		//
	short cItem;			// Item #									//
	short cPercent;			// Chance NPC will have item				//
};																		//
//////////////////////////////////////////////////////////////////////////