/********************************************************
 *  Class: Node											*
 *     By: Peter S. VanLund								*
 *   Desc: Defines a node and operations on it			*
 ********************************************************/

#include "NPC.h"
#include <vector>

using namespace std;

// Possible neighboring directions
enum Direction {NORTH,SOUTH,EAST,WEST};

class Node
{
	private:
		Node           *north,*south,*east,*west;	// Neighbors
		vector<Sound*> vecSound;					// Song list
		vector<Item*>  vecItem;						// Item list
		vector<Item*>  vecRequiredItem;				// Req item list
		vector<NPC*>   vecNPC;						// NPC list
		bool           visited;						// visited list
	public:
		// Constructor & Destructor
		Node();
		~Node();
		void  setNeighbor(Node*,Direction);	// Set neighbor
		Node* getNeighbor(Direction);		// Get neighbor
		void  addSound(Sound*);				// Add sound
		void  addItem(Item*);				// Add item
		Item* getItem();					// Pop item from list
		Item* getItem(int);					// Get item without removing
		void removeItem(int);				// Remove item
		void removeAllItems();				// Remove all items
		int getItemCount();					// Get item count
		void  addRequiredItem(Item*);		// Add req item
		void  addNPC(NPC*);					// Add NPC
		bool hasNPC();						// NPC test
		bool hasEnemy();					// Enemy test
		bool hasFriend();					// Friend test
		bool hasVendor();					// Vendor test
		bool hasLeprechaun();				// Leprechaun test
		NPC* getNPC();						// Get NPC
		void removeNPC();					// Remove NPC
		void visit();						// Visit node
		bool hasBeenVisited();				// Visit test
		bool hasSong(Sound *s);				// Sound test
		Sound* getRandomSong();				// Get random song
		Item* getRequiredItem(int);			// Get required item
		int getRequiredItemCount();			// Get number of req items
};

// Constructor
Node::Node()
{
	north = south = east = west = NULL;
	visited = false;
}
// Destructor
Node::~Node()
{
	while(vecNPC.size() > 0)
	{
		delete vecNPC[0];
		vecNPC.erase(vecNPC.begin());
	}
}

// setNeighbor: Sets neighbor to the d direction
void Node::setNeighbor(Node *n,Direction d)
{
	switch(d)
	{
		case NORTH: north = n; break;
		case SOUTH: south = n; break;
		case EAST:  east  = n; break;
		case WEST:  west  = n; break;
	}
}

// getNeighbor: Returns neighbor to the d direction
Node* Node::getNeighbor(Direction d)
{
	switch(d)
	{
		case NORTH: return north;
		case SOUTH: return south;
		case EAST:  return east;
		case WEST:  return west;
	}
	return NULL;
}

// Add sound & item
void Node::addSound(Sound *s){vecSound.push_back(s);}
void Node::addItem(Item *i){vecItem.push_back(i);}

// getItem: Pops item from list
Item* Node::getItem()
{
	if(vecItem.size()==0) return NULL;
	Item *i = vecItem[0];
	vecItem.erase(vecItem.begin());
	return i;
}

// getItem: Gets item without removal
Item* Node::getItem(int i)
{
	if(i>vecItem.size()-1) return NULL;
	return vecItem[i];
}

// removeItem: Remove item
void Node::removeItem(int i)
{
	if(i>vecItem.size()-1) return;
	vecItem.erase(vecItem.begin()+i);
}

// removeAllItems: Remove all items
void Node::removeAllItems()
{
	while(vecItem.size()>0)
		vecItem.erase(vecItem.begin());
}

// getItemCount: Get number of items
int Node::getItemCount(){return vecItem.size();}

// addRequiredItem: Add req item
void Node::addRequiredItem(Item *i){vecRequiredItem.push_back(i);}

// addNPC: Add NPC to node
void Node::addNPC(NPC *c)
{
	// Create copy of NPC for node
	NPC *newNPC = new NPC();
	*newNPC = *c;
	newNPC->randomize();
	vecNPC.push_back(newNPC);
}

// NPC tests
bool Node::hasNPC(){return vecNPC.size()>0;}
bool Node::hasEnemy(){return vecNPC.size()>0 && vecNPC[0]->getType()==ENEMY;}
bool Node::hasVendor(){return vecNPC.size()>0 && vecNPC[0]->getType()==VENDOR;}
bool Node::hasFriend(){return vecNPC.size()>0 && vecNPC[0]->getType()==FRIEND;}
bool Node::hasLeprechaun(){return vecNPC.size()>0 && vecNPC[0]->getType()==LEPRECHAUN;}

// getNPC: Get NPC from node
NPC* Node::getNPC()
{
	if(vecNPC.size()>0) return vecNPC[0];
	return NULL;
}
// removeNPC: Remove NPC
void Node::removeNPC()
{
	if(vecNPC.size()==0) return;
	delete vecNPC[0];
	vecNPC.erase(vecNPC.begin());
}

// Visit set & test
void Node::visit(){visited=true;}
bool Node::hasBeenVisited(){return visited;}

// hasSong: Sound test
bool Node::hasSong(Sound *s)
{
	for(int i = 0; i<vecSound.size(); i++)
		if(vecSound[i]==s) return true;
	return false;
}

// getRandomSong: Return a random song from list
Sound* Node::getRandomSong()
{
	if(vecSound.size()==0) return NULL;
	return vecSound[Random::getShort(0,vecSound.size()-1)];
}

// getRequiredItem: Get req item
Item* Node::getRequiredItem(int i)
{
	if(i>=vecRequiredItem.size()) return NULL;
	return vecRequiredItem[i];
}

// getRequiredItemCount: Get number of req items
int Node::getRequiredItemCount(){return vecRequiredItem.size();}
