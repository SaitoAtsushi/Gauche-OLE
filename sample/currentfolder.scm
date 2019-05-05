#!/usr/bin/env gosh
(use win.ole)
(use gauche.collection)

(define (ShowFolderFileList folderspec)
  (let* ((fso (make-ole "Scripting.FileSystemObject"))
         (f (fso 'GetFolder folderspec)))
    (for-each (^x (print (~ x 'Name))) (~ f 'files))))

(ShowFolderFileList ".")

(ole-release!)
