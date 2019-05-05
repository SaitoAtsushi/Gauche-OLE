
#include <string.h>
#include "gauche-ole.h"
#include "callback.h"

STDMETHODIMP_(ULONG) ICallback_AddRef(ICallback* p) {
  return ++(p->counter);
}

STDMETHODIMP_(ULONG) ICallback_Release(ICallback* p) {
  int counter = --(p->counter);
  if(counter==0) {
    free(p);
    p->proc = NULL;
  }
  return counter;
}

STDMETHODIMP ICallback_QueryInterface(ICallback* p, REFIID riid,
                                      PVOID* ppvObject) {
  if(memcmp(&riid, &IID_IUnknown, sizeof(riid))) {
    *ppvObject = p;
  } else if(memcmp(&riid, &IID_IDispatch, sizeof(riid))) {
    *ppvObject = p;
  } else {
    *ppvObject = NULL;
    return E_NOINTERFACE;
  }

  ICallback_AddRef(p);
  return S_OK;
}

STDMETHODIMP ICallback_GetTypeInfoCount(ICallback* p, UINT* pctinfo) {
  *pctinfo = 0;
  return NOERROR;
}

STDMETHODIMP ICallback_GetTypeInfo(ICallback* p, UINT iTInfo, LCID lcid,
                                   LPTYPEINFO* ppTInfo) {
  *ppTInfo = NULL;
  return iTInfo ? S_OK : DISP_E_BADINDEX;
}

STDMETHODIMP ICallback_GetIDsOfNames(ICallback* p, REFIID riid,
                                     LPOLESTR* rgszNames, UINT cNames,
                                     LCID lcid, DISPID* rgDispId) {
  return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP ICallback_Invoke(ICallback* p, DISPID dispID, REFIID riid,
                              LCID lcid, WORD wFlags, DISPPARAMS* pDispParams,
                              VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                              UINT* puArgErr) {
  if(dispID != DISPID_VALUE) return DISP_E_MEMBERNOTFOUND;
  ScmObj args = SCM_NIL;
  for(int i=0; i<pDispParams->cArgs; i++)
    args = SCM_OBJ(Scm_Cons(VariantToScmObj(&pDispParams->rgvarg[i], FALSE),
                            args));
  ScmObj retVal = Scm_ApplyRec(SCM_OBJ(p->proc), args);
  if(pVarResult) {
    VariantInit(pVarResult);
    ScmObjToVariant(retVal, pVarResult);
  }
  return S_OK;
}

struct ICallbackVtbl ICallbackVtblImpl = {
  ICallback_QueryInterface,
  ICallback_AddRef,
  ICallback_Release,
  ICallback_GetTypeInfoCount,
  ICallback_GetTypeInfo,
  ICallback_GetIDsOfNames,
  ICallback_Invoke
};

ICallback* MakeCallbackHandler(ScmProcedure* proc) {
  ICallback* pCallback = malloc(sizeof(ICallback));
  pCallback->counter=1;
  pCallback->proc=proc;
  pCallback->lpVtbl=&ICallbackVtblImpl;
  return pCallback;
}
