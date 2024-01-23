// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package objecthandling.actions;

import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.webui.CustomJavaAction;

public class getMemberValueByMemberName extends CustomJavaAction<java.lang.String>
{
	private final IMendixObject MxObject;
	private final java.lang.String AttributeName;

	public getMemberValueByMemberName(
		IContext context,
		IMendixObject _mxObject,
		java.lang.String _attributeName
	)
	{
		super(context);
		this.MxObject = _mxObject;
		this.AttributeName = _attributeName;
	}

	@java.lang.Override
	public java.lang.String executeAction() throws Exception
	{
		// BEGIN USER CODE
		
		return MxObject.getValue(getContext(), AttributeName);
		
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "getMemberValueByMemberName";
	}

	// BEGIN EXTRA CODE
	// END EXTRA CODE
}