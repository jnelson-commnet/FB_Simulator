# FB_Simulator
Simulates actionable events for inventory reporting.

Current Bugs - 3/15/17
	- find_timing_issues() in ForecastMain is getting tricked by its input.  Should be correctable by using the inventory results from add_inv_counter().
	- The timeline is coming out with poorly assigned GrandParents (though Parents might be working as intended).  Seems like the issue is most likely in find_demand_driver().

Bug found - 10/2/17
	- fix_uom() was added to ForecastMain.py in order to edit BOMs with UOMs not matching the default UOM of a part.  That seems to be working, but the UOM issue is also showing from WOs made from those BOMs.  The WO item qty is being used as is even if it does not match the default.  This is likely an issue with SOs and POs as well.  fix_uom() is currently only set up to work with BOMs.  It should be reworked to adjust issues across all orders as well.
