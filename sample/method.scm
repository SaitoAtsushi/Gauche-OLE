#!/usr/bin/env gosh
(use win.ole)

;; see docment of VARFLAGS about method type
;; http://msdn.microsoft.com/en-us/library/windows/desktop/ms221426(v=vs.85).aspx

(define request (make-ole "Microsoft.XMLHTTP"))
(let1 const_values (ole-methods request)
  (for-each (^x (format #t "~a : ~a~%" (car x) (cdr x))) const_values))
(ole-release!)
