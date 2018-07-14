#pragma once
#include "stdafx.h"

namespace NetOffice_ShimLoader
{
	/*
	Determines the Assembly Registration Mode
	*/
	enum RegisterMode
	{
		// Component is registered per Machine
		// Addin is registered per current User
		SystemComponentAndUserAddin = 0,

		// Component and Addin is registered per Machine
		System = 1,

		// Component and Addin is registered per current User
		// This cause issues when application is started with admin privileges
		User = 2,
	};
}
