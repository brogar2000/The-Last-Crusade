/********************************************************
 *  Class: NPC											*
 *     By: Peter S. VanLund								*
 *   Desc: Used to create non-player characters			*
 ********************************************************/

#include "Item.h"

// NPC types
enum NPCType {ENEMY, FRIEND, VENDOR, LEPRECHAUN};

class NPC
{
	private:
		string        name;							// Name
		NPCType       type;							// Type
		short         str, def, hp, maxHP;			// Actual strenght, defense, etc.
		short         mxs, mns, mxd, mnd, mxh, mnh;	// [Min,Max] [Str,Def]
		short         runPerc;						// Odds of running away
		Sound         *nameSound, *actionSound;		// Sounds
		vector<Item*> vecItem;						// Items
		vector<int>   vecItemOdds;					// Odds of having items
	public:
		// Constructors & Destructor
		NPC(string,short,short,short,short,short,short,short,short,Sound*,Sound*);
		NPC();
		~NPC();
		void randomize();			// Randomize NPC items & stats
		void addItem(Item*,int);	// Add item
		string  getName();			// Get name
		NPCType getType();			// Get type
		short   getStr();			// Get strength
		void    setStr(int);		// Set strength
		short   getDef();			// Get defense
		void    setDef(int);		// Set defense
		short   getHP();			// Get hitpoints
		void    setHP(int);			// Set hitpoints
		short   getRun();			// Get run %
		Sound*  getNameSound();		// Get name sound
		Sound*  getActionSound();	// Get action sound
		void    playNameSound();	// Play name sound
		void    playActionSound();	// Player action sound
		void    heal();				// Heal NPC
		void    subHP(int);			// Substract hitpoints
		bool    isDead();			// Death test
		Item*   getItem();			// Pop item off list
		Item*   getItem(int);		// Get item i (not removed from list)
		void    removeItem(int);	// Remove item i from list
		void    removeAllItems();	// Remove all items
		int     getItemCount();		// Get number of items
};

// Constructor: Given NPC data
NPC::NPC(string n, short t, short st1, short st2, short df1, short df2, short h1, short h2, short r, Sound *nam, Sound *act)
{
	name        = n;
	type        = (NPCType)t;
	mns         = st1;
	mxs         = st2;
	mnd         = df1;
	mxd         = df2;
	mnh         = h1;
	mxh         = h2;
	runPerc     = r;
	nameSound   = nam;
	actionSound = act;
}

// Empty blank Constructor & Destructor
NPC::NPC(){}
NPC::~NPC(){}

// randomize: Randomize NPC items & stats
void NPC::randomize()
{
	// Stats
	str = Random::getShort(mns,mxs);
	def = Random::getShort(mnd,mxd);
	hp = maxHP = Random::getShort(mnh,mxh);
	int i;
	// Items
	for(i = 0; i<vecItem.size(); i++)
	{
		if(!Random::checkOdds(vecItemOdds[i]))
		{
			vecItem.erase(vecItem.begin()+i);
			vecItemOdds.erase(vecItemOdds.begin()+i);
			i--;
		}
	}
}

// addItem: Adds item to NPC (i=Item*,o=odds)
void NPC::addItem(Item *i,int o)
{
	vecItem.push_back(i);
	vecItemOdds.push_back(o);
}

// NPC getter & setter methods
string NPC::getName(){return name;}
NPCType NPC::getType(){return type;}
short NPC::getStr(){return str;}
void NPC::setStr(int s){str=s;}
short NPC::getDef(){return def;}
void NPC::setDef(int d){def=d;}
short NPC::getHP(){return hp;}
void NPC::setHP(int h){hp=maxHP=h;}
short NPC::getRun(){return runPerc;}
Sound* NPC::getNameSound(){return nameSound;}
Sound* NPC::getActionSound(){return actionSound;}

// Sound playing methods
void NPC::playNameSound(){if(nameSound!=NULL) nameSound->playAndWait(type==FRIEND);}
void NPC::playActionSound(){if(actionSound!=NULL) actionSound->playAndWait(type==FRIEND);}

// Misc methods
void NPC::heal(){hp = maxHP;}
void NPC::subHP(int i){hp-=i;}
bool NPC::isDead(){return hp<=0;}

// getItem: Pops item off NPC item list
Item* NPC::getItem()
{
	if(vecItem.size()==0) return NULL;
	Item *i = vecItem[0];
	vecItem.erase(vecItem.begin());
	return i;
}

// getItem: Gets item off item list without removing it
Item* NPC::getItem(int i)
{
	if(i>vecItem.size()-1) return NULL;
	return vecItem[i];
}

// removeItem: Removes item from item list
void NPC::removeItem(int i)
{
	if(i>vecItem.size()-1) return;
	vecItem.erase(vecItem.begin()+i);
}

// removeAllItems: Removes all items from item list
void NPC::removeAllItems()
{
	while(vecItem.size()>0)
		vecItem.erase(vecItem.begin());
}

// getItemCount: Get number of items
int NPC::getItemCount(){return vecItem.size();}