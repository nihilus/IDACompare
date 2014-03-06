/*

 Sample how to integrate VB UI for IDA plugin
 http://sandsprite.com/CodeStuff/VB_Plugin_for_Olly.html

'Author: David Zimmer <dzzie@yahoo.com> - Copyright 2004
'Site:   http://sandsprite.com
'

NOTE: to build this project it is assumed you have an envirnoment variable 
named IDASDK set to point to the base SDK directory. this env var is used in
the C/C++ property tab, Preprocessor catagory. I have also had to harcode the paths
to my ida sdk lib files below..apparently pragma comment does not accept env vars

finally, the exports of this lib have been set to accept and return int64 types always
for addresses. this lets 32/64 bit handling be more standardized in the layers above.

'License:
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA


*/


//#define __EA64__  //create the plugin for the 32 bit, 64 bit capable IDA

#ifdef __EA64__
	#pragma comment(linker, "/out:./../Ida_Compare.p64")
	#pragma comment(lib, "D:\\idasdk65\\idasdk65\\lib\\x86_win_vc_64\\ida.lib")
#else
	#pragma comment(linker, "/out:./../Ida_Compare.plw")
	#pragma comment(lib, "D:\\idasdk65\\idasdk65\\lib\\x86_win_vc_32\\ida.lib")
#endif

#pragma warning(disable:4996) //may be unsafe function
#pragma warning(disable:4244) //conversion from '__int64' to 'ea_t', possible loss of data

#include <windows.h>  //define this before other headers or get errors 
#include <ida.hpp>
#include <idp.hpp>
#include <expr.hpp>
#include <bytes.hpp>
#include <loader.hpp>
#include <kernwin.hpp>
#include <name.hpp>
#include <auto.hpp>
#include <frame.hpp>
#include <dbg.hpp>
#include <area.hpp>

#undef strcpy
#undef sprintf


IDispatch        *IDisp;

int StartPlugin(void);
int idaapi init(void){ return PLUGIN_OK; }
void idaapi run(int arg){ StartPlugin(); }

void idaapi term(void)
{
	try{
		if(IDisp){
			IDisp->Release();
			CoUninitialize();
			IDisp = NULL;
		}
	}
	catch(...){};
	
}

char comment[] = "idacompare";
char help[] ="idacompare";
char wanted_name[] = "IDA Compare";
char wanted_hotkey[] = "Alt-0";

//Plugin Descriptor Block
plugin_t PLUGIN =
{
  IDP_INTERFACE_VERSION,
  0,                    // plugin flags
  init,                 // initialize
  term,                 // terminate. this pointer may be NULL.
  run,                  // invoke plugin
  comment,              // long comment about the plugin (status line or hint)
  help,                 // multiline help about the plugin
  wanted_name,          // the preferred short name of the plugin
  wanted_hotkey         // the preferred hotkey to run the plugin
};





int StartPlugin(){

    //Create an instance of our VB COM object, and execute
	//one of its methods so that it will load up and show a UI
	//for us, then it uses our other exports to access olly plugin API
	//methods

	CLSID      clsid;
	HRESULT	   hr;
    LPOLESTR   p = OLESTR("IDACompare.CPlugin");

    hr = CoInitialize(NULL);

	 hr = CLSIDFromProgID( p , &clsid);
	 if( hr != S_OK  ){
		 MessageBox(0,"Failed to get Clsid from string\n","",0);
		 return 0;
	 }

	 // create an instance and get IDispatch pointer
	 hr =  CoCreateInstance( clsid,
							 NULL,
							 CLSCTX_INPROC_SERVER,
							 IID_IDispatch  ,
							 (void**) &IDisp
						   );

	 if ( hr != S_OK )
	 {
	   MessageBox(0,"CoCreate failed","",0);
	   return 0;
	 }

	 OLECHAR *sMethodName = OLESTR("DoPluginAction");
	 DISPID  dispid; // long integer containing the dispatch ID

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDS failed","",0);
		return 0;
	 }

	 DISPPARAMS dispparams;
	 VARIANTARG vararg[1]; //function takes one argument
	 VARIANT    retVal;

	 VariantInit(&vararg[0]);
	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 0;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);

	 return 0;
}





//Export API for the VB app to call and access IDA API data
//_________________________________________________________________
void __stdcall Refresh   (void)      { refresh_idaview();      }
void __stdcall Setname( ea_t addr, const char* name){ set_name((ea_t)addr, name); }
void __stdcall MessageUI(char *m){ msg(m);}
int  __stdcall FuncIndex(__int64 addr){ return get_func_num((ea_t)addr); }
void __stdcall FuncName(__int64 addr, char *buf, size_t bufsize){ get_func_name((ea_t)addr, buf, bufsize);}
int  __stdcall GetBytes(__int64 offset, void *buf, int length){ return get_many_bytes((ea_t)offset, buf, length);}
int __stdcall FilePath(char *buf){ return get_input_file_path(buf,255); }
int __stdcall NumFuncs  (void){ return get_func_qty(); }


//retrieves function names and jump labels 
void __stdcall GetName(__int64 offset, char* buf, int bufsize){

	get_true_name( BADADDR, (ea_t)offset, buf, bufsize );

	if(strlen(buf) == 0){
		func_t* f = get_func((ea_t)offset);
		for(int i=0; i < f->llabelqty; i++){
			if( f->llabels[i].ea == offset ){
				int sz = strlen(f->llabels[i].name);
				if(sz < bufsize) strcpy(buf,f->llabels[i].name);
				return;
			}
		}
	}

}

bool FuncIndexOk(int n, bool warn = true){ //ida will crash if out of bounds..
	
	char buf[100];
	if(n < 0 || n >= NumFuncs()){
		if(warn){
			sprintf(&buf[0], "Invalid FunctionStart(%x)", n);
			MessageBoxA(0,buf,"PLW",0);
		}
		return false;
	}
	
	return true;
}

__int64 __stdcall Addx64(__int64 base, unsigned int val){
	return base + val;
}

__int64 __stdcall Subx64(__int64 v0, __int64 v1){
	return v0 - v1;
}

__int64 __stdcall FunctionStart(int n){
	if(!FuncIndexOk(n)) return 0;
	func_t *clsFx = getn_func(n);
	return (__int64)clsFx->startEA;
}

__int64 __stdcall FunctionEnd(int n){
	if(!FuncIndexOk(n)) return 0;
	func_t *clsFx = getn_func(n);
	return (__int64)clsFx->endEA;
}

int __stdcall GetAsm(__int64 addr, char* buf, int bufLen){

    flags_t flags;                                                       
    int sLen=0;

    flags = getFlags(addr);                        
    if(isCode(flags)) {                            
        generate_disasm_line((ea_t)addr, buf, bufLen, GENDSM_MULTI_LINE );
        sLen = tag_remove(buf, buf, bufLen);  
    }

	return sLen;

}

