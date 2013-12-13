/*
Created By Erik L. Eidt.
This file is part of JEB.

Copyright (c) 2011, 2012, Hewlett-Packard Development Company, L.P.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.
    * Neither the name of Hewlett-Packard nor the
      names of its contributors may be used to endorse or promote products
      derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL 
HEWLETT-PACKARD DEVELOPMENT COMPANY, L.P. OR THE CONTRIBUTORS 
BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

package com.hp.JEB;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.DispatchAddon;
import com.jacob.com.Variant;

public class Application {
	private static boolean			itsIsComAttached	= false;
	private static boolean			itsWasNew			= false;
	public static ActiveXComponent	xa					= null;

	public static Workbook			ActiveWorkbook		= null;
	public static Worksheet			ActiveSheet			= null;

	public static boolean Initialize ( boolean createIfNotFound ) {
		boolean ans = false;
		try {
			ComThreadInit ();
			itsWasNew = false;
			xa = ActiveXComponent
					.connectToActiveInstance ( "Excel.Application.14" );
			if ( xa == null ) {
				xa = ActiveXComponent
						.connectToActiveInstance ( "Excel.Application.13" );
				if ( xa == null ) {
					xa = ActiveXComponent
							.connectToActiveInstance ( "Excel.Application.13" );
					if ( xa == null ) {
						xa = ActiveXComponent
								.connectToActiveInstance ( "Excel.Application.12" );
						if ( xa == null ) {
							xa = ActiveXComponent
									.connectToActiveInstance ( "Excel.Application" );
							if ( xa == null && createIfNotFound ) {
								xa = new ActiveXComponent ( "Excel.Application" );
								if ( xa != null ) {
									itsWasNew = true;
									Dispatch.put ( xa, "Visible", Constants.kTrue );
								}
							}
						}
					}
				}
			}
		} catch ( Exception e ) {}

		if ( xa == null ) {
			ComThreadRelease ();
		} else {
			ActiveWorkbook = GetActiveWorkbook ();
			ActiveSheet = GetActiveWorksheet ();
			ans = true;
		}
		return ans;
	}
	
	public static boolean connected () {
		return xa != null;
	}

	public static void Release () {
		Quit ();
		xa = null;
		if ( itsIsComAttached ) {
			ComThread.Release ();
			itsIsComAttached = false;
		}
	}

	public static void Quit () {
		if ( itsWasNew ) {
			if ( xa != null )
				xa.invoke ( "Quit", Constants.kEmptyA );
			itsWasNew = false;
		}
	}

	public static Variant getProperty ( int dispID ) {
		return Dispatch.get ( xa, dispID );
	}

	// ======================================================================
	public static Worksheet GetActiveWorksheet () {
		// Dispatch d = xa.getProperty("ActiveSheet").toDispatch();
		Dispatch d = getProperty ( 0x133 ).toDispatch ();
		if ( d == null )
			return null;
		return new Worksheet ( d );
	}

	public static Workbook GetActiveWorkbook () {
		// Dispatch d = xa.getProperty("Activeworkbook").toDispatch();
		Dispatch d = getProperty ( 0x134 ).toDispatch ();
		if ( d == null )
			return null;
		return new Workbook ( d );
	}

	public static Workbooks Workbooks () {
		// Dispatch d = xa.getProperty("Workbooks").toDispatch();
		Dispatch d = getProperty ( 0x23c ).toDispatch ();
		if ( d == null )
			return null;
		return new Workbooks ( d );
	}

	public static Workbook Workbooks ( String wkbName ) {
		return Workbooks ().Item ( wkbName );
	}

	public static Workbook Workbooks ( int num ) {
		return Workbooks ().Item ( num );
	}

	public static void StatusBar ( String msg ) {
		// Dispatch.put(xa, "StatusBar", msg);
		Dispatch.put ( xa, 0x182, msg );
	}

	public static int ScreenUpdating () {
		// return Dispatch.get(xa, "ScreenUpdating").toInt();
		Variant v = Dispatch.get ( xa, 0x17e );
		v.changeType ( Variant.VariantInt );
		return v.getInt (); // .toInt();
	}

	public static void SetScreenUpdating ( int su ) {
		// Dispatch.put(xa, "ScreenUpdating", su);
		Dispatch.put ( xa, 0x17e, su );
		// lang.ErrMsg("ScreenUpdating =" + su);
	}

	public static Range Selection () {
		// Dispatch d = Dispatch.get(xa, "Selection").toDispatch();
		Variant v = Dispatch.get ( xa, 0x93 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public static Range Intersect ( Range r1, Range r2 ) {
		// Dispatch d = Dispatch.get2p(xa, "Intersect", r1, r2).toDispatch();
		Variant v = DispatchAddon.get2p ( xa, 0x2fe, r1, r2 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	/*
	 * public static String Run(String vbName) { // Variant v =
	 * Dispatch.call(xa, "Run", vbName); Variant v = Dispatch.call(xa, 0x103,
	 * vbName); //return Dispatch.get(this, "Name").getString(); return
	 * v.getString(); }
	 * 
	 * public static Variant Run1p(String vbName, Object param) { // Variant v =
	 * Dispatch.call(xa, "Run", vbName, param); return Dispatch.call(xa, 0x103,
	 * vbName, param); //return Dispatch.get(this, "Name").getString(); }
	 * 
	 * public static Variant Run2p(String vbName, Object p1, Object p2) { //
	 * Variant v = Dispatch.call(xa, "Run", vbName, param1, param2); return
	 * Dispatch.call(xa, 0x103, vbName, p1, p2); //return Dispatch.get(this,
	 * "Name").getString(); }
	 * 
	 * public static Variant Run4p(String vbName, Object p1, Object p2, Object
	 * p3, Object p4) { // Variant v = Dispatch.call(xa, "Run", vbName, param1,
	 * param2); return Dispatch.call(xa, 0x103, vbName, p1, p2, p3, p4);
	 * //return Dispatch.get(this, "Name").getString(); }
	 */

	// ----------------------------------------------------------------------

	private static void ComThreadInit () {
		ComThread.InitSTA ();
		itsIsComAttached = true;
	}

	private static void ComThreadRelease () {
		if ( itsIsComAttached ) {
			itsIsComAttached = false;
			ComThread.Release ();
		}
	}
}
