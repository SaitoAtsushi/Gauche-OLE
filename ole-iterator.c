
#define COBJMACROS

#include <windows.h>
#include <objbase.h>
#include <oaidl.h>
#include "gauche-ole.h"
#include "ole-iterator.h"

static ScmClass *default_cpl[] = {
  SCM_CLASS_STATIC_PTR(Scm_OLEIteratorClass),
  SCM_CLASS_STATIC_PTR(Scm_TopClass), NULL
};

static void ole_iterator_print(ScmObj obj, ScmPort *out, ScmWriteContext *ctx) {
  if(SCM_OLE(obj)->pDispatch) {
    Scm_Printf(out, "#<ole-iterator %p>", SCM_OLE_ITERATOR(obj)->pDispatch);
  } else {
    Scm_Printf(out, "#<ole-iterator released>");
  }
}

SCM_DEFINE_BASE_CLASS(Scm_OLEIteratorClass, ScmOLEIterator,
                      ole_iterator_print, NULL, NULL, NULL,
                      default_cpl);

IEnumVARIANT* make_ole_iterator(IDispatch *pDispatch) {
  if(!pDispatch)
    Scm_Error("OLE object was released.");
  
  DISPID dispid = DISPID_NEWENUM;
  DISPPARAMS dispParams;
  dispParams.cArgs = 0;
  dispParams.rgvarg = NULL;
  dispParams.cNamedArgs = 0;
  dispParams.rgdispidNamedArgs = NULL;

  VARIANT VarResult;
  VariantInit(&VarResult);
  HRESULT hr = IDispatch_Invoke(pDispatch, dispid, &IID_NULL,
                                LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET,
                                &dispParams, &VarResult, NULL, NULL);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);
  if(VarResult.vt != VT_UNKNOWN)
    Scm_Error("Internal error in make-ole-iterator");

  IEnumVARIANT *pEnumVariant;
  hr = IUnknown_QueryInterface(VarResult.punkVal, &IID_IEnumVARIANT,
                               (void **)&pEnumVariant);
  if(FAILED(hr))
    Scm_RaiseOLECondition(hr);
  return pEnumVariant;
}

ScmObj ole_iterator_next(IEnumVARIANT* pEnumVARIANT) {
  if(!pEnumVARIANT)
    Scm_Error("OLE object was released.");

  VARIANT Var;
  VariantInit(&Var);
  HRESULT hr = pEnumVARIANT->lpVtbl->Next(pEnumVARIANT, 1, &Var, NULL);
  return hr==S_OK ? VariantToScmObj(&Var, TRUE): SCM_EOF;
}
 
void ole_iterator_release(ScmObj obj) {
  if(!SCM_OLE_ITERATOR_P(obj))
    Scm_Error("ole-iterator object required, but got %S", obj);
  if(SCM_OLE_ITERATOR(obj)->pDispatch!=NULL) {
    IUnknown_Release(SCM_OLE_ITERATOR(obj)->pDispatch);
    SCM_OLE_ITERATOR(obj)->pDispatch = NULL;
  }
}

void Scm_OLE_Iterator_Finalizer(ScmObj obj, void* data) {
  ole_iterator_release(obj);
}

ScmObj OLE_ITERATOR_HANDLE_BOX(IEnumVARIANT* pDispatch) {
  ScmOLEIterator* iter = SCM_OLE_ITERATOR(SCM_OBJ(SCM_NEW(ScmOLEIterator)));
  SCM_SET_CLASS(iter, SCM_CLASS_OLE_ITERATOR) ;
  SCM_OLE_ITERATOR(iter)->pDispatch = pDispatch;
  Scm_RegisterFinalizer(SCM_OBJ(iter), Scm_OLE_Iterator_Finalizer, NULL);
  return SCM_OBJ(iter);
}
