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

package com.jacob.com;

public class DispatchAddon {
	private static int []	zArgErr	= {};

	public static boolean isAttached ( Dispatch theOneInQuestion ) {
		return theOneInQuestion.isAttached ();
	}

	private static void throwIfUnattachedDispatch ( Dispatch theOneInQuestion ) {
		if ( theOneInQuestion == null ) {
			throw new IllegalArgumentException (
					"Can't pass in null Dispatch object" );
		} else if ( theOneInQuestion.isAttached () ) {
			return;
		} else {
			throw new IllegalStateException (
					"Dispatch not hooked to windows memory" );
		}
	}

	public static Variant get1p ( Dispatch dispatchTarget, String name, Object oArg ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		return Dispatch.invokev ( dispatchTarget, name, Dispatch.Get, new Variant [] { VariantUtilities.objectToVariant ( oArg ) }, zArgErr );
	}

	public static Variant get1p ( Dispatch dispatchTarget, int dispID, Object oArg ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		return Dispatch.invokev ( dispatchTarget, dispID, Dispatch.Get, new Variant [] { VariantUtilities.objectToVariant ( oArg ) }, zArgErr );
	}

	public static Variant get2p ( Dispatch dispatchTarget, String name, Object oArg1, Object oArg2 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ) };
		return Dispatch.invokev ( dispatchTarget, name, Dispatch.Get, va, zArgErr );
	}

	public static Variant get2p ( Dispatch dispatchTarget, int dispID, Object oArg1, Object oArg2 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ) };
		return Dispatch.invokev ( dispatchTarget, dispID, Dispatch.Get, va, zArgErr );
	}

	public static Variant get3p ( Dispatch dispatchTarget, String name, Object oArg1, Object oArg2, Object oArg3 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ), VariantUtilities.objectToVariant ( oArg3 ) };
		return Dispatch.invokev ( dispatchTarget, name, Dispatch.Get, va, zArgErr );
	}

	public static Variant get3p ( Dispatch dispatchTarget, int dispID, Object oArg1, Object oArg2, Object oArg3 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ), VariantUtilities.objectToVariant ( oArg3 ) };
		return Dispatch.invokev ( dispatchTarget, dispID, Dispatch.Get, va, zArgErr );
	}

	public static Variant get4p ( Dispatch dispatchTarget, String name, Object oArg1, Object oArg2, Object oArg3, Object oArg4 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ),
				VariantUtilities.objectToVariant ( oArg3 ), VariantUtilities.objectToVariant ( oArg4 ), };
		return Dispatch.invokev ( dispatchTarget, name, Dispatch.Get, va, zArgErr );
	}

	public static Variant get4p ( Dispatch dispatchTarget, int dispID, Object oArg1, Object oArg2, Object oArg3, Object oArg4 ) {
		throwIfUnattachedDispatch ( dispatchTarget );
		Variant [] va = new Variant [] { VariantUtilities.objectToVariant ( oArg1 ), VariantUtilities.objectToVariant ( oArg2 ),
				VariantUtilities.objectToVariant ( oArg3 ), VariantUtilities.objectToVariant ( oArg4 ), };
		return Dispatch.invokev ( dispatchTarget, dispID, Dispatch.Get, va, zArgErr );
	}
}
