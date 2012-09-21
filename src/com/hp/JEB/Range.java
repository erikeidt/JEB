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

public class Range extends Dispatch {
	static String	dbgLast;

	public Range ( Dispatch d ) {
		super ( d );
	}

	public String Address () {
		// return Dispatch.get(this, "Address").toString();
		return Dispatch.get ( this, 0xec ).toString ();
	}

	public String Address ( boolean a, boolean b ) {
		// return Dispatch.get2p(this, "Address", a, b).toString();
		return DispatchAddon.get2p ( this, 0xec, a, b ).toString ();
	}

	public String Value () {
		String x;
		// Variant v = Dispatch.get(this, "Value");
		Variant v = Dispatch.get ( this, 0x6 );
		if ( v == null || v.isNull () )
			return "";
		if ( v.getvt () == Variant.VariantEmpty ) {
			dbgLast = "";
			return "";
		}
		switch ( v.getvt () ) {
		case Variant.VariantEmpty :
			return "";
		case Variant.VariantDouble :
			x = java.lang.Double.toString ( v.getDouble () ); // .toDouble()
			if ( x.endsWith ( ".0" ) )
				x = x.substring ( 0, x.length () - 2 );
			dbgLast = x;
			return x;

		}
		x = v.getString ();
		dbgLast = x;
		return x;
		//return v.getString();
	}

	public void SetValue ( String val ) {
		// Dispatch.put(this, "Value", val);		
		Dispatch.put ( this, 0x6, val );
	}

	public void SetValue ( int val ) {
		// Dispatch.put(this, "Value", val);		
		Dispatch.put ( this, 0x6, val );
	}

	public void SetFormula ( String f ) {
		Dispatch.put ( this, "Formula", f );
	}

	public void Select () {
		// Dispatch.call (this, "Select");
		Dispatch.call ( this, 0xeb );
	}

	public void AutoFit () {
		Dispatch.call ( this, "AutoFit" );
	}

	public void Merge () {
		Dispatch.call ( this, "Merge" );
	}

	public Worksheet Parent () {
		// Dispatch d = Dispatch.get(this, "Parent").toDispatch();
		Variant v = Dispatch.get ( this, 0x96 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Worksheet ( d );
	}

	public int Column () {
		// return Dispatch.get(this, "Column").toInt();		
		return Dispatch.get ( this, 0xf0 ).getInt (); // .toInt();		
	}

	public int Count () {
		// return Dispatch.get(this, "Count").toInt();
		return Dispatch.get ( this, 0x76 ).getInt (); // .toInt();
	}

	public Range Columns () {
		// Dispatch d = Dispatch.get(this, "Columns").toDispatch();
		Variant v = Dispatch.get ( this, 0xf1 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Columns ( int num ) {
		// Dispatch d = Dispatch.get1p(this, "Columns", num).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0xf1, num );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public int Row () {
		// return Dispatch.get(this, "Row").toInt();		
		return Dispatch.get ( this, 0x101 ).getInt (); // .toInt();		
	}

	public Range Rows () {
		// Dispatch d = Dispatch.get(this, "Rows").toDispatch();
		Variant v = Dispatch.get ( this, 0x102 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Rows ( int num ) {
		// Dispatch d = Dispatch.get1p(this, "Rows", num).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0x102, num );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Areas Areas () {
		// Dispatch d = Dispatch.get(this, "Areas").toDispatch();
		Variant v = Dispatch.get ( this, 0x238 );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Areas ( d );
	}

	public Range Areas ( int num ) {
		return Areas ().Item ( num );
		/*
		Dispatch d = Dispatch.get1p(this, "Area", num).toDispatch();
		return new Range(d);
		*/
	}

	public Range Offset ( int row, int col ) {
		// Dispatch d = Dispatch.get2p(this, "Offset", row, col).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0xfe, row, col );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Offset ( int row ) {
		// Dispatch d = Dispatch.get1p(this, "Offset", row).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0xfe, row );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Resize ( int row, int col ) {
		// Dispatch d = Dispatch.get2p(this, "Resize", row, col).toDispatch();
		Variant v = DispatchAddon.get2p ( this, 0x100, row, col );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range Resize ( int row ) {
		// Dispatch d = Dispatch.get1p(this, "Resize", row).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0x100, row );
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

	public Range End ( int xlV ) {
		// Dispatch d = Dispatch.get1p(this, "End", xlV).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0x1f4, xlV );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public Range SpecialCells ( int xlV ) {
		// Dispatch d = Dispatch.get1p(this, "SpecialCells", xlV).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0x19a, xlV );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Range ( d );
	}

	public int InteriorColor () {
		// Dispatch d = Dispatch.get(this, "Interior").toDispatch();
		Dispatch d = Dispatch.get ( this, 0x81 ).toDispatch ();
		Variant v = Dispatch.get ( d, "Color" );
		v.changeType ( Variant.VariantInt );
		return v.getInt (); // .toInt();
	}

	public int FontColor () {
		Dispatch d = Dispatch.get ( this, "Font" ).toDispatch ();
		if ( d == null )
			return 0;
		Variant v = Dispatch.get ( d, "Color" );
		v.changeType ( Variant.VariantInt );
		return v.getInt (); // .toInt();		
	}

	public void SetInteriorColor ( int clr ) {
		// Dispatch d = Dispatch.get(this, "Interior").toDispatch();
		Dispatch d = Dispatch.get ( this, 0x81 ).toDispatch ();
		Dispatch.put ( d, "Color", clr );
	}

	public Font Font () {
		Variant v = Dispatch.get ( this, "Font" );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Font ( d );
	}

	public void SetColumnWidth ( int w ) {
		Dispatch.put ( this, "ColumnWidth", w );
	}

	public void SetReportColor ( int clr ) {
		Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.SetReportColor", this, clr );
	}

	public void SetDefaultReportFormat () {
		Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.SetDefaultReportFormat" );
	}

	/*
	public String GetQuickRow(int row, int col1, int col2) {
		// Variant v = Dispatch.call(xa, "Run", "QuickRDA.JavaCallBacks.GetQuickRow", rng, row, col1, col2);
		return Dispatch.call(Application.xa, 0x103, "QuickRDA.JavaCallBacks.GetQuickRow", this, row, col1, col2).getString();
		//return Dispatch.get(this, "Name").getString();
	}
	*/

	public Validation Validation () {
		Variant v = Dispatch.get ( this, "Validation" );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new Validation ( d );
	}

	public String GetQuickTab ( int row1, int row2, int col1, int col2, int vis ) {
		// Variant v = Dispatch.call(xa, "Run", "QuickRDA.JavaCallBacks.GetQuickTab", rng, row1, row2, col1, col2);
		Variant v = Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.GetQuickTab", this, row1, row2, col1, col2, vis );
		if ( v == null || v.isNull () )
			return null;
		return v.getString ();
		//return Dispatch.get(this, "Name").getString();
	}

	public void SetValidation ( Workbook wkb, int targetColNum, int sourceColIndex, int sourceLength ) {
		Dispatch.call ( Application.xa, 0x103, "QuickRDA.JavaCallBacks.SetValidationTarget", this, wkb, targetColNum, sourceColIndex, sourceLength );
	}
}
