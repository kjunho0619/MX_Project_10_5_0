// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package oidc.actions;

import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;
import oidc.proxies.Claim;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.thirdparty.org.json.JSONObject;

public class CreateClaimsString extends CustomJavaAction<java.lang.String>
{
	/** @deprecated use com.mendix.utils.ListUtils.map(claimsList, com.mendix.systemwideinterfaces.core.IEntityProxy::getMendixObject) instead. */
	@java.lang.Deprecated(forRemoval = true)
	private final java.util.List<IMendixObject> __claimsList;
	private final java.util.List<oidc.proxies.Claim> claimsList;

	public CreateClaimsString(
		IContext context,
		java.util.List<IMendixObject> _claimsList
	)
	{
		super(context);
		this.__claimsList = _claimsList;
		this.claimsList = java.util.Optional.ofNullable(_claimsList)
			.orElse(java.util.Collections.emptyList())
			.stream()
			.map(claimsListElement -> oidc.proxies.Claim.initialize(getContext(), claimsListElement))
			.collect(java.util.stream.Collectors.toList());
	}

	@java.lang.Override
	public java.lang.String executeAction() throws Exception
	{
		// BEGIN USER CODE
		JSONObject result = new JSONObject();
		JSONObject userinfo = new JSONObject();
		for (Claim claim : claimsList) {
			userinfo.put(claim.getValue(), JSONObject.NULL);
		}
		result.put("userinfo", userinfo);
		return result.toString();
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "CreateClaimsString";
	}

	// BEGIN EXTRA CODE
	// END EXTRA CODE
}
