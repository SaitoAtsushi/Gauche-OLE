
typedef struct ScmOLEIteratorRec {
  SCM_HEADER;
  IEnumVARIANT* pDispatch;
} ScmOLEIterator;

SCM_CLASS_DECL(Scm_OLEIteratorClass);
#define SCM_CLASS_OLE_ITERATOR (&Scm_OLEIteratorClass)
#define SCM_OLE_ITERATOR(obj)  ((ScmOLEIterator*)obj)
#define SCM_OLE_ITERATOR_P(obj) SCM_XTYPEP(obj, SCM_CLASS_OLE_ITERATOR)
#define OLE_ITERATOR_HANDLE_UNBOX(obj) (SCM_OLE_ITERATOR(obj)->pDispatch)
ScmObj OLE_ITERATOR_HANDLE_BOX(IEnumVARIANT* pDispatch);

IEnumVARIANT* make_ole_iterator(IDispatch*);
ScmObj ole_iterator_next(IEnumVARIANT*);
void ole_iterator_release(ScmObj);
