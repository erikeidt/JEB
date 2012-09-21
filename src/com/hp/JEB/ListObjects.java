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

import com.jacob.com.DispatchAddon;
import com.jacob.com.Variant;
import com.jacob.com.Dispatch;

public class ListObjects extends Dispatch {
	public ListObjects ( Dispatch d ) {
		super ( d );
	}

	public int Count () {
		// Variant v = Dispatch.get(this, "Count");
		Variant v = Dispatch.get ( this, 0x76 );
		return v.getInt (); // .toInt();
	}

	public ListObject Item ( int num ) {
		// Dispatch d = Dispatch.get1p(this, "Item", num).toDispatch();
		Variant v = DispatchAddon.get1p ( this, 0xaa, num );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new ListObject ( d );
	}

	public ListObject Add ( int srcType, Range src, int hdrType ) {
		Variant v = Dispatch.call ( this, "Add", srcType, src, Variant.VT_MISSING, hdrType );
		if ( v == null || v.isNull () )
			return null;
		Dispatch d = v.toDispatch ();
		if ( d == null )
			return null;
		return new ListObject ( d );
	}
}
