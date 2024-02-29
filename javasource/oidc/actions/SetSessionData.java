// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package oidc.actions;

import java.lang.reflect.Method;
import com.mendix.core.Core;
import com.mendix.core.conf.RuntimeVersion;
import com.mendix.m2ee.api.IMxRuntimeRequest;
import com.mendix.m2ee.api.IMxRuntimeResponse;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;
import system.proxies.TokenInformation;
import com.mendix.systemwideinterfaces.core.ISession;
import com.mendix.systemwideinterfaces.core.IUser;

public class SetSessionData extends CustomJavaAction<java.lang.Boolean>
{
	private final java.lang.String Username;

	public SetSessionData(
		IContext context,
		java.lang.String _username
	)
	{
		super(context);
		this.Username = _username;
	}

	@java.lang.Override
	public java.lang.Boolean executeAction() throws Exception
	{
		// BEGIN USER CODE
		IContext ctx = getContext();
		IUser user = Core.getUser(ctx, this.Username);
		if (user != null) {
			ISession session = Core.initializeSession(user, null);
			// get the response object
			if (ctx.getRuntimeResponse().isPresent()) {
				IMxRuntimeResponse res = ctx.getRuntimeResponse().get();
				setCookies(res, session); // set xassessionid and xasid cookies
				return true;
			}

		}

		return false;

		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "SetSessionData";
	}

	// BEGIN EXTRA CODE
	private void setCookies(IMxRuntimeResponse response, ISession session) throws Exception{
		response.addCookie(Core.getConfiguration().getSessionIdCookieName(), session.getId().toString(), "/", "", -1, true,true);
		response.addCookie("XASID", "0." + Core.getXASId(), "/", "", -1, true);
	}
	
	// END EXTRA CODE
}
