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

import com.jacob.com.Dispatch;
import com.jacob.com.DispatchAddon;
import com.jacob.com.Variant;

public class Workbook extends Dispatch {
	public Workbook ( Dispatch d ) {
		super ( d );
	}

	public String Name () {
		//return Dispatch.get(this, "Name").getString();
		return Dispatch.get ( this, 0x6e ).getString ();
	}

	public Worksheet ActiveSheet () {
		// Dispatch d = Dispatch.get(this, "ActiveSheet").toDispatch();
		Variant v = Dispatch.get ( this, 0x133 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch (); // must be the same as Application.ActiveSheet
		if ( d == null )
			return null;
		return new Worksheet ( d );
	}

	public Worksheets Worksheets () {
		// Dispatch d = Dispatch.get(this, "Worksheets").toDispatch();
		Variant v = Dispatch.get ( this, 0x1ee );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch (); // must be the same as Application.Worksheets
		if ( d == null )
			return null;
		return new Worksheets ( d );
	}

	public Worksheet Worksheets ( int num ) {
		return Worksheets ().Item ( num );
	}

	public Worksheet Worksheets ( String wkbName ) {
		return Worksheets ().Item ( wkbName );
	}

	public String Path () {
		// Variant v = Dispatch.get(this, "Path");
		Variant v = Dispatch.get ( this, 0x123 ); // must be the same as Application.Path
		if ( v == null || v.isNull () )
			return null;
		return v.getString ();
	}

	public Names Names () {
		// Dispatch d = Dispatch.get(this, "Names").toDispatch();
		Variant v = Dispatch.get ( this, 0x1ba );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch (); // must be the same as Application.Path
		if ( d == null )
			return null;
		return new Names ( d );
	}

	public Name Names ( String n ) {
		Names nms = Names ();
		return nms.Item ( n );
		// Dispatch d = Dispatch.get1p(this, "Names", n).toDispatch();
		// return new Name(d);
	}

	public void Activate () {
		// Dispatch.call(this, "Activate");
		Dispatch.call ( this, 0x130 );
	}

	public void SaveAs ( String name ) {
		Dispatch.call ( this, "SaveAs", name );
		// Dispatch.call(this, 0x130, name);		
	}

	public void Close ( boolean saveChanges, String fileName, boolean doRouting ) {
		// Dispatch.call (this, "Close", saveChanges, fileName, doRouting);
		Dispatch.call ( this, 0x115, saveChanges, fileName, doRouting );
	}

	public Range HasInfoTable () {
		Variant v = Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.HasInfoTable", this );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null || !DispatchAddon.isAttached ( d ) )
			return null;
		return new Range ( d );
	}

	public void SetDropDownSources ( String tab ) {
		Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.SetDropDownSources", this, tab );
	}

}
