
struct ICallbackVtbl;

typedef struct ICallbackRec {
  struct ICallbackVtbl *lpVtbl;
  int counter;
  ScmProcedure* proc;
} ICallback;

struct ICallbackVtbl {
  STDMETHOD(QueryInterface)(ICallback*, REFIID,PVOID*);
  STDMETHOD_(ULONG,AddRef)(ICallback*);
  STDMETHOD_(ULONG,Release)(ICallback*);
  STDMETHOD(GetTypeInfoCount)(ICallback*, UINT*);
  STDMETHOD(GetTypeInfo)(ICallback*, UINT,LCID,LPTYPEINFO*);
  STDMETHOD(GetIDsOfNames)(ICallback*, REFIID,LPOLESTR*,UINT,LCID,DISPID*);
  STDMETHOD(Invoke)(ICallback*, DISPID,REFIID,LCID,WORD,DISPPARAMS*,VARIANT*,EXCEPINFO*,UINT*);
};

ICallback* MakeCallbackHandler(ScmProcedure*);
