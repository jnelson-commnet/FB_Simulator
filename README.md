# FB_Simulator
Simulates actionable events for inventory reporting.

Current Bugs - 3/15/17
	- find_timing_issues() in ForecastMain is getting tricked by its input.  Should be correctable by using the inventory results from add_inv_counter().
	- The timeline is coming out with poorly assigned GrandParents (though Parents might be working as intended).  Seems like the issue is most likely in find_demand_driver().
