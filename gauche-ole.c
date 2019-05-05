
#define COBJMACROS

#include <windows.h>
#include <objbase.h>
#include <oaidl.h>
#include <string.h>
#include "gauche-ole.h"
#include "callback.h"
#include "weak_ex.h"

#define SCM_ISA_OLE(obj) SCM_ISA((obj), &Scm_OLEClass)
#define SCM_ISA_OLECollection(obj) SCM_ISA((obj), &Scm_OLECollectionClass)

static struct {
  ScmWeakVector* objects;
  ScmInternalMutex mutex;
  ScmSmallInt index;
} active_ole_objects = { NULL };

#define OLE_TABLE_INITIAL_SIZE 1024

void Scm_OLE_Release(ScmObj obj);

void ole_release_all(void) {
  ScmWeakVector* objects = active_ole_objects.objects;

  SCM_INTERNAL_MUTEX_LOCK(active_ole_objects.mutex);
  for(int i=0; i<active_ole_objects.index; i++) {
    ScmObj obj = Scm_WeakVectorRef(objects, i, SCM_FALSE);
    if(SCM_ISA_OLE(obj)) Scm_OLE_Release(SCM_OBJ(obj));
  }
  Scm_WeakVectorRealloc(objects, OLE_TABLE_INITIAL_SIZE);
  active_ole_objects.index = 0;

  SCM_INTERNAL_MUTEX_UNLOCK(active_ole_objects.mutex);
}

static void register_active_ole(ScmOLE* obj) {
  ScmWeakVector* objects = active_ole_objects.objects;
  SCM_INTERNAL_MUTEX_LOCK(active_ole_objects.mutex);

  if(objects->size == active_ole_objects.index) {
    GC_gcollect();
    for(int i; i<objects->size; i++) {
      ScmObj obj = Scm_WeakVectorRef(objects, i, SCM_FALSE);
      if(SCM_ISA_OLE(obj) && SCM_OLE(obj)->pDispatch==NULL)
        Scm_WeakVectorSet(objects, i, SCM_FALSE);
    }

    active_ole_objects.index = Scm_WeakVectorSqueeze(objects);
    if(objects->size == active_ole_objects.index)
      Scm_WeakVectorRealloc(objects, 2*objects->size);
  }

  Scm_WeakVectorSet(active_ole_objects.objects,
                    active_ole_objects.index++,
                    SCM_OBJ(obj));
  SCM_INTERNAL_MUTEX_UNLOCK(active_ole_objects.mutex);
}

ScmModule* this_module;

static void ole_assert(ScmObj obj) {
  if(!SCM_ISA_OLE(obj))
    Scm_Error("ole object required, but got %S", obj);
}

static void ole_collection_assert(ScmObj obj) {
  if(!SCM_ISA_OLECollection(obj))
    Scm_Error("ole-collection object required, but got %S", obj);
}

static ScmClass *default_cpl[] = {
  SCM_CLASS_STATIC_PTR(Scm_OLEClass),
  SCM_CLASS_STATIC_PTR(Scm_TopClass), NULL
};

static void ole_print(ScmObj obj, ScmPort *out, ScmWriteContext *ctx) {
  if(SCM_OLE(obj)->pDispatch) {
    Scm_Printf(out, "#<ole %p>", SCM_OLE(obj)->pDispatch);
  } else {
    Scm_Printf(out, "#<ole released>");
  }
}

SCM_DEFINE_BASE_CLASS(Scm_OLEClass, ScmOLE, ole_print, NULL, NULL, NULL, default_cpl);

static ScmClass *ole_collection_cpl[] = {
  SCM_CLASS_STATIC_PTR(Scm_OLEClass),
  SCM_CLASS_STATIC_PTR(Scm_CollectionClass),  
  SCM_CLASS_STATIC_PTR(Scm_TopClass), NULL
};

static void ole_collection_print(ScmObj obj, ScmPort *out, ScmWriteContext *ctx) {
  if(SCM_OLE(obj)->pDispatch) {
    Scm_Printf(out, "#<ole-collection %p>", SCM_OLE_COLLECTION(obj)->pDispatch);
  } else {
    Scm_Printf(out, "#<ole-collection released>");
  }
}

SCM_DEFINE_BASE_CLASS(Scm_OLECollectionClass, ScmOLE,
                      ole_collection_print, NULL, NULL, NULL,
                      ole_collection_cpl);

void Scm_RaiseOLECondition(HRESULT number) {
  Scm_RaiseCondition(SCM_SYMBOL_VALUE("win.ole", "<ole-condition>"),
                     "number", Scm_MakeInteger(number),
                     SCM_RAISE_CONDITION_MESSAGE,
                     "OLE error was thrown. result code = %d", number);
}

static OLECHAR* ScmStringToOLEString(ScmString* str) {
  return SysAllocString(Scm_MBS2WCS(SCM_STRING_START(str)));
}

static HRESULT getID(IDispatch* pDispatch, ScmSymbol* method, DISPID* dispID) {
  OLECHAR* method_name = ScmStringToOLEString(SCM_SYMBOL_NAME(method));
  HRESULT hr =
    IDispatch_GetIDsOfNames(pDispatch, &IID_NULL, &method_name, 1, LOCALE_USER_DEFAULT, dispID);
  SysFreeString(method_name);
  return hr;
}

static void ole_safe_release(IDispatch* pDispatch) {
  if(pDispatch) IDispatch_Release(pDispatch);
}

static void Scm_OLE_Finalizer(ScmObj obj, void *data) {
  ole_safe_release(SCM_OLE(obj)->pDispatch);
}

static ScmObj make_OLE(IDispatch* pDispatch) {
  ScmObj obj = SCM_OBJ(SCM_NEW(ScmOLE));
  SCM_OLE_COLLECTION(obj)->pDispatch = pDispatch;
  if(pDispatch!=NULL) {
    DISPID dispid;
    WCHAR* mname = L"_NewEnum";
    HRESULT hr =
      IDispatch_GetIDsOfNames(pDispatch, &IID_NULL, &mname, 1, LOCALE_USER_DEFAULT, &dispid);
    if(FAILED(hr) || dispid!=DISPID_NEWENUM)
      SCM_SET_CLASS(obj, SCM_CLASS_OLE);
    else
      SCM_SET_CLASS(obj, SCM_CLASS_OLE_COLLECTION);

    Scm_RegisterFinalizer(obj, Scm_OLE_Finalizer, NULL);
    register_active_ole(SCM_OLE(obj));
  } else
    SCM_SET_CLASS(obj, SCM_CLASS_OLE);

  return SCM_OBJ(obj);
}

static ScmString* OLEStringToScmString(OLECHAR* str) {
  return SCM_STRING(SCM_MAKE_STR_IMMUTABLE(str==NULL ? "" : Scm_WCS2MBS(str)));
}


//#define ITypeLib_GetTypeInfo(p, a, b) ((p)->lpVtbl->GetTypeInfo(p, a, b))
#define ITypeLib_GetTypeAttr(p, a) ((p)->lpVtbl->GetTypeAttr((p), (a)))
//#define ITypeLib_GetTypeInfoCount(p) ((p)->lpVtbl->GetTypeInfoCount(p))
//#define ITypeLib_Release(p) ((p)->lpVtbl->Release(p))


ScmObj VariantToScmObj(VARIANT* var, int flag) {
  static ScmObj array = NULL;
  static ScmObj shape = NULL;

  if(var==NULL)
    return SCM_UNDEFINED;

  ScmObj retVal;
  switch(var->vt) {
  case VT_NULL:
    retVal = SCM_NIL;
    break;
  case VT_EMPTY:
    retVal = SCM_UNDEFINED;
    break;
  case VT_I8:
    retVal = Scm_MakeInteger64(var->llVal);
    break;
  case VT_I4:
    retVal = Scm_MakeInteger(var->lVal);
    break;
  case VT_UI1:
    retVal = Scm_MakeIntegerU(var->bVal);
    break;
  case VT_I2:
    retVal = Scm_MakeInteger(var->iVal);
    break;
  case VT_BOOL:
    retVal = SCM_MAKE_BOOL(var->boolVal==VARIANT_TRUE);
    break;
  case VT_BSTR:
    retVal = SCM_OBJ(OLEStringToScmString(var->bstrVal));
    if(flag) {
      SysFreeString(var->bstrVal);
      var->bstrVal=NULL;
    }
    break;
  case VT_INT:
    retVal = Scm_MakeInteger(var->intVal);
    break;
  case VT_R8:
    retVal = Scm_MakeFlonum(var->dblVal);
    break;
  case VT_R4:
    retVal = Scm_MakeFlonum(var->fltVal);
    break;
  case VT_DISPATCH:
    if(var->pdispVal != NULL) IDispatch_AddRef(var->pdispVal);
    retVal = make_OLE(var->pdispVal);
    break;
  case (VT_ARRAY | VT_VARIANT) :
    if(array == NULL) {
      ScmSymbol* sym_array = SCM_SYMBOL(SCM_INTERN("array"));
      ScmSymbol* sym_shape = SCM_SYMBOL(SCM_INTERN("shape"));
      ScmModule* array_module = SCM_FIND_MODULE("gauche.array", FALSE);
      array = Scm_SymbolValue(array_module, sym_array);
      shape = Scm_SymbolValue(array_module, sym_shape);
    }
    SAFEARRAY* sa = var->parray;
    USHORT dim = sa->cDims;
    ScmObj shape_item = SCM_NIL;

    long* index = malloc((1+dim)* sizeof(long));
    long* min_bound = malloc(dim* sizeof(long));
    long* max_bound = malloc(dim* sizeof(long));

    for(int i=0; i<dim; i++)
      min_bound[dim-i-1] = sa->rgsabound[i].lLbound;

    for(int i=0; i<dim; i++) {
      max_bound[dim-i-1] =
        sa->rgsabound[i].lLbound + sa->rgsabound[i].cElements;
      index[dim-i-1] = max_bound[dim-i-1] - 1;
    }

    for(int i=dim-1; i>=0; i--) {
      shape_item = Scm_Cons(Scm_MakeInteger(max_bound[i]), shape_item);
      shape_item = Scm_Cons(Scm_MakeInteger(min_bound[i]), shape_item);
    }

    ScmObj retShape = Scm_ApplyRec(shape, shape_item);

    ScmObj al = SCM_NIL;
    while(1) {
      VARIANT val;
      VariantInit(&val);
      HRESULT hr = SafeArrayGetElement(sa, index, &val);
      if(FAILED(hr))
        Scm_RaiseOLECondition(hr);
      ScmObj obj = VariantToScmObj(&val, TRUE);
      al = Scm_Cons(obj, al);
      index[dim-1]--;
      int j;
      for(j=dim-1; index[j]<min_bound[j] && j>=0; j--) {
        index[j] = max_bound[j]-1;
        if(j) index[j-1]--;
      }
      if(j<0) break;
    }
    al = Scm_Cons(retShape, al);
    retVal = Scm_ApplyRec(array, al);
    SafeArrayDestroy(sa);
    free(index);
    free(min_bound);
    free(max_bound);
    break;
  default:
    Scm_Error("Not support type of VARIANT: %d", var->vt);
    retVal = SCM_FALSE;
  }

  if(flag) VariantClear(var);
  return retVal;
}

int ScmObjToVariant(ScmObj obj, VARIANT* pResult) {
  if(SCM_NULLP(obj)) {
    pResult->vt = VT_NULL;
  } else if(SCM_UNDEFINEDP(obj)) {
    pResult->vt = VT_EMPTY;
  } else if(SCM_STRINGP(obj)) {
    pResult->bstrVal = ScmStringToOLEString(SCM_STRING(obj));
    pResult->vt = VT_BSTR;
  } else if(SCM_BOOLP(obj)) {
    pResult->boolVal = SCM_TRUEP(obj) ? VARIANT_TRUE : VARIANT_FALSE;
    pResult->vt = VT_BOOL;
  } else if(SCM_INTP(obj)) {
    pResult->intVal = SCM_INT_VALUE(obj);
    pResult->vt = VT_INT;
  } else if(SCM_FLONUMP(obj)) {
    pResult->dblVal = SCM_FLONUM_VALUE(obj);
    pResult->vt = VT_R8;
  } else if(SCM_PROCEDUREP(obj)) {
    pResult->pdispVal = (IDispatch*) MakeCallbackHandler(SCM_PROCEDURE(obj));
    pResult->vt = VT_DISPATCH;
  } else if(SCM_OLE_P(obj)) {
    pResult->pdispVal = SCM_OLE(obj)->pDispatch;
    pResult->vt = VT_DISPATCH;
  } else {
    Scm_Error("Unable to convert to VARIANT type: %S", obj);
  }
  return 1;
}

ScmObj Scm_OLE_Const(ScmObj obj) {
  ole_assert(obj);
  IDispatch* pDispatch = SCM_OLE(obj)->pDispatch;
    
  ScmObj retVal = SCM_NIL;

  if(!pDispatch)
    Scm_Error("OLE object was released.");

  ITypeInfo *pTInfo;
  HRESULT hr = IDispatch_GetTypeInfo(pDispatch, 0, LOCALE_USER_DEFAULT, &pTInfo);

  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  ITypeLib *pTLib;
  UINT index;
  hr = ITypeInfo_GetContainingTypeLib(pTInfo, &pTLib, &index);

  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  ITypeInfo_Release(pTInfo);

  // int count = ITypeLib_GetTypeInfoCount(pTLib);
  for(int i=0; i<index; i++) {
    hr = ITypeLib_GetTypeInfo(pTLib, i, &pTInfo);
    if(FAILED(hr)) {
      Scm_RaiseOLECondition(hr);
    }
    TYPEATTR *pTAttr;
    ITypeLib_GetTypeAttr(pTInfo, &pTAttr);
    if(FAILED(hr)) {
      ITypeLib_Release(pTInfo);
      Scm_RaiseOLECondition(hr);
    }

    for(int iVar=0; iVar<(pTAttr->cVars); iVar++) {
      VARDESC *pVarDesc;
      hr = ITypeInfo_GetVarDesc(pTInfo, iVar, &pVarDesc);
      if(FAILED(hr)) continue;
      if(pVarDesc->varkind == VAR_CONST &&
         !(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
                                  VARFLAG_FRESTRICTED |
                                  VARFLAG_FNONBROWSABLE))) {
        BSTR bstr;
        UINT len;
        hr = ITypeInfo_GetNames(pTInfo, pVarDesc->memid, &bstr, 1, &len);
        if(FAILED(hr) || len == 0 || !bstr) {
          continue;
        }
        
        retVal = Scm_Cons(Scm_Cons(SCM_OBJ(Scm_Intern(OLEStringToScmString(bstr))),
                                   VariantToScmObj(pVarDesc->lpvarValue, TRUE)),
                          retVal);

        SysFreeString(bstr);        
      }
      ITypeInfo_ReleaseVarDesc(pTInfo, pVarDesc);
    }

    ITypeInfo_ReleaseTypeAttr(pTInfo, pTAttr);
    ITypeInfo_Release(pTInfo);
  }
  return retVal;
}

ScmObj Scm_OLE_Methods(ScmObj obj) {
  ole_assert(obj);
  IDispatch* pDispatch = SCM_OLE(obj)->pDispatch;

  ScmObj retVal = SCM_NIL;

  if(!pDispatch)
    Scm_Error("OLE object was released.");

  ITypeInfo *pTInfo;
  HRESULT hr = IDispatch_GetTypeInfo(pDispatch, 0, LOCALE_USER_DEFAULT, &pTInfo);

  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  ITypeLib *pTLib;
  UINT index;
  hr = ITypeInfo_GetContainingTypeLib(pTInfo, &pTLib, &index);

  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  ITypeInfo_Release(pTInfo);

  // int count = ITypeLib_GetTypeInfoCount(pTLib);
  int i=index;

  hr = ITypeLib_GetTypeInfo(pTLib, i, &pTInfo);
  if(FAILED(hr)) {
    Scm_RaiseOLECondition(hr);
  }
  TYPEATTR *pTAttr;
  ITypeLib_GetTypeAttr(pTInfo, &pTAttr);
  if(FAILED(hr)) {
    ITypeLib_Release(pTInfo);
    Scm_RaiseOLECondition(hr);
  }

  for(int iVar=0; iVar<(pTAttr->cFuncs); iVar++) {
    FUNCDESC  *pFuncDesc;
    hr = ITypeInfo_GetFuncDesc(pTInfo, iVar, &pFuncDesc);
    if(FAILED(hr)) {
      Scm_RaiseOLECondition(hr);
    }
    if(!(pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED|
                                  FUNCFLAG_FHIDDEN|
                                  FUNCFLAG_FNONBROWSABLE
                                  ))) {
      BSTR bstr;
      UINT len;
      hr = ITypeInfo_GetNames(pTInfo, pFuncDesc->memid, &bstr, 1, &len);
      if(FAILED(hr) || len == 0 || !bstr) continue;
        
      retVal = Scm_Cons(Scm_Cons(SCM_OBJ(Scm_Intern(OLEStringToScmString(bstr))),
                                 Scm_MakeInteger(pFuncDesc->invkind)),
                        retVal);

      SysFreeString(bstr);        
    }
    ITypeInfo_ReleaseFuncDesc(pTInfo, pFuncDesc);
  }
  ITypeInfo_ReleaseTypeAttr(pTInfo, pTAttr);
  ITypeInfo_Release(pTInfo);

  return retVal;
}

void Scm_OLE_Release(ScmObj obj) {
  ole_assert(obj);
  ole_safe_release(SCM_OLE(obj)->pDispatch);
  SCM_OLE(obj)->pDispatch = NULL;
  Scm_UnregisterFinalizer(obj);
}

ScmObj Scm_Dispatch(ScmString* progID) {
  CLSID id;

  HRESULT hr = CLSIDFromProgID(ScmStringToOLEString(progID), &id);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  IDispatch* pDispatch;
  hr = CoCreateInstance(&id, NULL,
                        CLSCTX_ALL,
                        &IID_IDispatch,
                        (void**)&pDispatch);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  return make_OLE(pDispatch);
}

ScmObj Scm_Connect(ScmString* progID) {
  CLSID id;

  HRESULT hr = CLSIDFromProgID(ScmStringToOLEString(progID), &id);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);

  IUnknown *pUnknown;
  hr = GetActiveObject(&id, 0, &pUnknown);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);
  
  IDispatch *pDispatch;
  hr = IUnknown_QueryInterface(pUnknown, &IID_IDispatch, (void**)&pDispatch);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);
  ole_safe_release((IDispatch*)pUnknown);

  return make_OLE(pDispatch);
}

ScmObj Scm_Invoke(ScmObj obj, ScmSymbol* method, ScmObj args, WORD flag) {
  ole_assert(obj);
  IDispatch* pDispatch = SCM_OLE(obj)->pDispatch;

  if(!pDispatch)
    Scm_Error("OLE object was released.");

  DISPID dispID;
  HRESULT hr = getID(pDispatch, method, &dispID);
  if(FAILED(hr)) {
    Scm_RaiseOLECondition(hr);
  }

  int args_len = (int)Scm_Length(SCM_OBJ(args));
  if(args_len < 0) {
    Scm_Error("Argument list is not proper", args);
  }

  VARIANTARG *pvArgs;
  if(args_len==0) {
    pvArgs=NULL;
  }
  else {
    pvArgs = malloc(sizeof(VARIANTARG)*args_len);
    if(pvArgs==NULL) {
      Scm_Error("Failed to memory allocation");
    }

    ScmObj p=SCM_OBJ(args);
    for(int i=args_len-1; !SCM_NULLP(p); p=SCM_CDR(p), i--){
      VariantInit(&pvArgs[i]);
      ScmObjToVariant(SCM_CAR(p), &pvArgs[i]);
    }
  }
  DISPPARAMS dispParams;
  dispParams.rgvarg = pvArgs;
  dispParams.rgdispidNamedArgs = NULL;
  dispParams.cArgs = args_len;
  dispParams.cNamedArgs = 0;

  VARIANT vResult;
  VariantInit(&vResult);
  hr = IDispatch_Invoke(pDispatch,
                        dispID, &IID_NULL,
                        LOCALE_USER_DEFAULT,
                        flag,
                        &dispParams, &vResult,
                        NULL, NULL);

  for(int i=0; i<args_len; i++)
    VariantClear(&pvArgs[i]);

  free(pvArgs);

  if(FAILED(hr)) {
    Scm_RaiseOLECondition(hr);
  }

  ScmObj result = VariantToScmObj(&vResult, TRUE);
  
  return result;
}

ScmObj Scm_CallMethod(ScmObj obj, ScmSymbol* method, ScmObj args) {
  return Scm_Invoke(obj, method, args, DISPATCH_METHOD);
}

ScmObj Scm_GetProp(ScmObj obj, ScmSymbol* prop, ScmObj args) {
  return Scm_Invoke(obj, prop, args, DISPATCH_PROPERTYGET);
}

extern ScmObj Scm_PutProp(ScmObj obj, ScmSymbol* prop, ScmObj var) {
  ole_assert(obj);
  IDispatch* pDispatch = SCM_OLE(obj)->pDispatch;

  if(!pDispatch)
    Scm_Error("OLE object was released.");

  DISPID id;
  DISPID dispIDNamedArgs[1] = { DISPID_PROPERTYPUT };

  HRESULT hr = getID(pDispatch, prop, &id);
  if(FAILED(hr)) {
    Scm_RaiseOLECondition(hr);
  }

  VARIANTARG vArgs;
  VariantInit(&vArgs);
  ScmObjToVariant(var, &vArgs);

  DISPPARAMS dispParams;
  dispParams.rgvarg = &vArgs;
  dispParams.rgdispidNamedArgs = dispIDNamedArgs;
  dispParams.cArgs = 1;
  dispParams.cNamedArgs = 1;

  hr = IDispatch_Invoke(pDispatch,
                        id, &IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_PROPERTYPUT,
                        &dispParams, NULL,
                        NULL, NULL);
  VariantClear(&vArgs);
  return FAILED(hr) ? SCM_FALSE : SCM_TRUE;
}

/*
 * Module initialization function.
 */
extern void Scm_Init_olelib(ScmModule*);

void Scm_Init_ole(void) {
  /* Register this DSO to Gauche */
  SCM_INIT_EXTENSION(ole);

  /* Create the module if it doesn't exist yet. */
  this_module = SCM_MODULE(SCM_FIND_MODULE("win.ole", TRUE));

  /* Register stub-generated procedures */
  Scm_Init_olelib(this_module);

  Scm_InitStaticClass(&Scm_OLEClass, "<ole>", this_module, NULL, 0);
  Scm_InitStaticClass(&Scm_OLECollectionClass, "<ole-collection>", this_module, NULL, 0);
  CoInitialize(NULL);

  SCM_INTERNAL_MUTEX_INIT(active_ole_objects.mutex);
  active_ole_objects.objects =
    SCM_WEAK_VECTOR(Scm_MakeWeakVector(OLE_TABLE_INITIAL_SIZE));
  active_ole_objects.index = 0;
}
