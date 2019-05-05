
#include "gauche-ole.h"
#include "weak_ex.h"

ScmObj Scm_WeakVectorRealloc(ScmWeakVector* v, ScmSmallInt new_size) {
  void* new_area = GC_realloc(v->pointers, new_size);
  if(new_area == NULL) return SCM_FALSE;
  ScmObj* p = (ScmObj*) new_area;
  for (ScmSmallInt i=v->size; i<new_size; i++) p[i] = SCM_FALSE;
  v->pointers = new_area;
  v->size = new_size;
  return SCM_TRUE;
}

ScmSmallInt Scm_WeakVectorSqueeze(ScmWeakVector* v) {
  ScmSmallInt j=0;
  for(ScmSmallInt i=0; i<v->size; i++) {
    ScmObj obj = Scm_WeakVectorRef(v, i, SCM_FALSE);
    if(!SCM_FALSEP(obj)) {
      if(i==j) j++;
      else Scm_WeakVectorSet(v, j++, obj);
    }
  }
  return j;
}
