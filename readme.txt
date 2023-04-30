

Note this is still in dev --> not end user functional yet.

Extra References:
	vbDevKit.dll   -  http://sandsprite.com/tools.php?id=3  (open source)
	spSubclass2.dll - https://github.com/dzzie/libs/tree/master/Subclass2

 External Treeview based project explorer addin for vb6 IDE
 
 I want a more suitable treeview based project explorer for the vb6 IDE 
 
 ColinE66 has a great owner drawn one
    https://www.vbforums.com/showthread.php?890617-Add-In-Large-Project-Organiser-(alternative-Project-Explorer)-No-sub-classing!&highlight=
	
 I wanted to take a stab at creating a treeview based one.
 
 I hate writing addins, so instead of developing this purely as a plug-in, its a "plug-out".
 
 There is a small IPC addin that sits in the IDE and accepts commands.
 The external treeview project explorer is its own process which just interacts with the stub.
 
 Its a little more complex in some ways, but it makes debugging more complex logic quicker in the end.
 (is my theory anyway)
 
 Currently:
	- can mirror a vb6 project treeview 
	- only supports a single project, no project groups (rare for me)
	- allows you to arbitrarily regroup nodes and add new folders (w/ drag drop)
	- auto synced with IDE events adding/renaming/removing files 
	- can save and restore trees to disk
	
	- reloaded trees will diff against current IDE files
	    - add files its missing from IDE
	    - mark files removed from the IDE  
	    - keep your own groupings, node expanded/collapsed state etc
		
	- logic is in place to add files/folders to treeview by browse or drag/drop
	    - this is not yet wired into adding files to the IDE yet 
 
	- form to filter search all nodes 
	
	- removing files from the treeview not supported, IDE VBComponents.Remove is broken
	
        - will eventually navigate IDE to display file on double click of node..