/*
 * ole.h
 */
#ifndef GAUCHE_OLE_H
#define GAUCHE_OLE_H

#include <gauche.h>
#include <gauche/extend.h>

SCM_DECL_BEGIN

typedef struct ScmOLERec {
  SCM_HEADER;
  IDispatch* pDispatch;
} ScmOLE;

SCM_CLASS_DECL(Scm_OLEClass);
#define SCM_CLASS_OLE (&Scm_OLEClass)
#define SCM_OLE(obj)  ((ScmOLE*)obj)
#define SCM_OLE_P(obj) SCM_XTYPEP(obj, SCM_CLASS_OLE)
#define OLE_HANDLE_UNBOX(obj) (SCM_OLE(obj)->pDispatch)
#define OLE_HANDLE_BOX(pDispatch) Scm_MakeOLE(pDispatch)

SCM_CLASS_DECL(Scm_OLECollectionClass);
#define SCM_CLASS_OLE_COLLECTION (&Scm_OLECollectionClass)
#define SCM_OLE_COLLECTION(obj)  ((ScmOLE*)obj)
#define SCM_OLE_COLLECTION_P(obj) SCM_XTYPEP(obj, SCM_CLASS_OLE_COLLECTION)

extern ScmObj Scm_Dispatch(ScmString*);
extern ScmObj Scm_Connect(ScmString*);
extern ScmObj Scm_CallMethod(ScmObj, ScmSymbol*, ScmObj);
extern ScmObj Scm_PutProp(ScmObj, ScmSymbol*, ScmObj);
extern ScmObj Scm_GetProp(ScmObj, ScmSymbol*, ScmObj);
extern ScmObj Scm_GetItem(ScmObj, unsigned int);
extern void Scm_OLE_Release(ScmObj);
extern ScmObj Scm_OLE_Const(ScmObj);
extern ScmObj Scm_OLE_Methods(ScmObj);

ScmObj VariantToScmObj(VARIANT*, int);
int ScmObjToVariant(ScmObj, VARIANT*);
void Scm_OLE_Release(ScmObj);
void ole_release_all(void);
void Scm_RaiseOLECondition(HRESULT);
SCM_DECL_END

#endif  /* GAUCHE_OLE_H */
