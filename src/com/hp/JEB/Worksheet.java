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

public class Worksheet extends Dispatch {
	public Worksheet ( Dispatch d ) {
		super ( d );
	}

	public String Name () {
		//return Dispatch.get(this, "Name").getString();
		return Dispatch.get ( this, 0x6e ).getString ();
	}

	public Workbook Parent () {
		// Dispatch d = Dispatch.get(this, "Parent").toDispatch();
		Variant v = Dispatch.get ( this, 0x96 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Workbook ( d );
	}

	public Range Range ( String addrString ) {
		// Dispatch d = Dispatch.get1p(this, "Range", addrString).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0xc5, addrString );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Range ( int row, int col ) {
		// Dispatch d = Dispatch.get2p(this, "Range", row, col).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0xc5, row, col );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Range ( Range r1, Range r2 ) {
		// Dispatch d = Dispatch.get2p(this, "Range", r1, r2).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0xc5, r1, r2 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Range ( String r1, String r2 ) {
		// Dispatch d = Dispatch.get2p(this, "Range", r1, r2).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0xc5, r1, r2 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public void Activate () {
		// Dispatch.call(this, "Activate");
		Dispatch.call ( this, 0x130 );
	}

	public Range Cells () {
		// Dispatch d = Dispatch.get2p(this, "Cells", row, col).toDispatch();
		Variant v = Dispatch.get ( this, 0xee );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Cells ( int row, int col ) {
		// Dispatch d = Dispatch.get2p(this, "Cells", row, col).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0xee, row, col );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Rows ( int row ) {
		// Dispatch d = Dispatch.get1p(this, "Rows",  row).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0x102, row );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Columns ( int row ) {
		// Dispatch d = Dispatch.get1p(this, "Columns",  row).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0xf1, row );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public int Visible () {
		// return Dispatch.get(this, "Visible").toInt();
		return Dispatch.get ( this, 0x22e ).getInt (); // .toInt();
	}

	public ListObjects ListObjects () {
		// Dispatch d = Dispatch.get(this, "ListObjects").toDispatch();
		Variant v = Dispatch.get ( this, 0x8d3 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new ListObjects ( d );
	}

	public Range GetSourceUnitTable () {
		// Variant v = Dispatch.call(xa, "Run", "QuickRDA.JavaCallBacks.GetSourceUnitTable", this);
		Variant v = Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.GetSourceUnitTable", this );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null || !DispatchAddon.isAttached ( d ) )
			return null;
		//return Dispatch.get(this, "Name").getString();
		return new Range ( d );
	}

	/*
	public Range GetMetamodelTable() {
		// Variant v = Dispatch.call(xa, "Run", "QuickRDA.JavaCallBacks.GetSourceUnitTable", this);
		Dispatch d = Dispatch.call(Application.xa, 0x103, "QuickRDA.JavaCallBacks.GetMetamodelTable", this).toDispatch();
		if ( d == null || !d.isAttached() )
			return null;
		//return Dispatch.get(this, "Name").getString();
		return new Range(d);
	}
	*/

	public Range FindTableRangeOnSheet () {
		// Variant v = Dispatch.call(xa, "Run", "QuickRDA.JavaCallBacks.FindTableRangeOnSheet", this);
		Variant v = Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.FindTableRangeOnSheet", this );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		//return Dispatch.get(this, "Name").getString();
		return new Range ( d );
	}

	public int SetReport ( int asz, String repName, String itemPlural, String filePath, String wkbName ) {
		Variant v = Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.SetReport", this, asz, repName, itemPlural, filePath, wkbName );
		return v.getInt ();

	}
}