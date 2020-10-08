OfficeAddin

Usage:
	1. Create a new class (using Action1.vb as a template)
	
		Rename the class to the method you want to perform
		Update the constants with information for the method
		
	2. Create the new class in CmdBars.vb.  Decide what section it should go in.
		CreateCommonMenu
		CreateCommonToolBar
		CreateAppMenu
		CreateAppToolbar
	
Notes:
	Update AddFooter to display a dialog box with some options.
	Need to address Excel and other Apps that do not automatically update
	footer information.  Catch OnSave and update if custom property indicating footer set.