#pragma once
#include "stdafx.h"

/*
Determines the Assembly Registration Mode
*/
enum RegisterMode
{
	// Component and Addin is registered per Machine
	System = 0,

	// Component and Addin is registered per current User
	// This cause issues when application is started with admin privileges
	User = 1,

	// Component is registered per Machine
	// Addin is registered per current User
	SystemComponentAndUserAddin = 2
};
