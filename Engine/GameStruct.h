//
// Player Save File Data Structures
//
struct PLAYERITEM
{
	char piName[20];
	short piType;
	short piValue;
	char piNameSound[50];
	char piActionSound[50];
};

struct PLAYERHEADER
{
	short pMap;
	short pNode;
	short pGold;
	short pMaxHP;
	short pHP;
	short pStr;
	short pDef;
	PLAYERITEM pWeapon;
	PLAYERITEM pArmor;
	short pPotionCount;
	short pSpecialCount;
};

struct PLAYERNODEHEADER
{
	bool pnVisited;
	short pnNPCCount;
	short pnItemCount;
};

struct PLAYERNPCHEADER
{
	short npcIndex;
	short npcStr;
	short npcDef;
	short npcHP;
	short npcItemCount;
};