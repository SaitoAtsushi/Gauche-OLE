#!/usr/bin/env gosh
(use win.ole)

(define request (make-ole "Microsoft.XMLHTTP"))
(let1 const_values (ole-const request)
  (for-each (^x (format #t "~a : ~a~%" (car x) (cdr x))) const_values))
(ole-release!)
