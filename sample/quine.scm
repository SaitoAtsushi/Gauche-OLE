#!/usr/bin/env gosh
(use win.ole)

(define (main args)
  (define fs (make-ole "Scripting.FileSystemObject"))
  (define file (fs 'OpenTextFile (car args)))
  (display (file 'ReadAll))
  (file 'close)
  (ole-release!)
)
