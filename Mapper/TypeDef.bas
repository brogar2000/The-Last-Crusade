Attribute VB_Name = "TypeDef"
' Type definitions for all map IO

Type MAPHEADER
    mStart As Integer        ' Start node
    mEnd As Integer          ' End node
    mNodeCount As Integer    ' Number of nodes
    mNPCCount As Integer     ' Number of characters
    mSoundCount As Integer   ' Number of sounds (events | dialogue | music)
    mItemCount As Integer    ' Number of items
End Type

Type NODEHEADER
    nIndex As Integer        ' Index in this VB program
    nNorth As Integer        ' Node to north
    nSouth As Integer        ' Node to south
    nEast As Integer         ' Node to east
    nWest As Integer         ' Node to west
    nMusicCount As Integer   ' # of songs for node
    nItemCount As Integer    ' # of items for node
    nNPCCount As Integer     ' # of NPCs for node
    nReqItemCount As Integer ' # of required items
    nImage As String * 20    ' Image for node
End Type

Type NODESOUND
    nSound As Integer
End Type

Type NODEITEM
    nItem As Integer
    nPercent As Integer
End Type

Type NODENPC
    nNPC As Integer
    nPercent As Integer
End Type

Type NODEREQITEM
    nReqItem As Integer
End Type

Type SOUNDDATA
    sName As String * 20     ' MP3 filename
    sSpatial As Boolean      ' Boolean for spatial
    sNode As Integer         ' Node source
    sXCoord As Integer       ' Spatial x coord
    sYCoord As Integer       ' Spatial y coord
End Type

Type ITEMDATA
    iName As String * 20     ' Item name
    iType As Integer         ' Weapon | Armor | Potion | Gold | Special
    iValue As Integer        ' Item value
    iNameSound As Integer    ' Name sound
    iActionSound As Integer  ' Action sound
End Type

Type NPCDATA
    cName As String * 20     ' NPC name
    cType As Integer         ' Enemy | Friend | Vendor | Leprechaun
    cStrMin As Integer       ' Strength min
    cStrMax As Integer       ' Strength max
    cDefMin As Integer       ' Defense min
    cDefMax As Integer       ' Defense max
    cHPMin As Integer        ' HP min
    cHPMax As Integer        ' HP max
    cRunPerc As Integer      ' Run %
    cNameSound As Integer    ' Name sound
    cActionSound As Integer  ' Action sound
    cItemCount As Integer    ' # of items
End Type

Type NPCITEM
    cItem As Integer
    cPercent As Integer
End Type
