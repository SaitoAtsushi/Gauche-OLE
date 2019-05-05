#!/usr/bin/env gosh

;; Write into Kingsoft Spreadsheet.
;; If you wish use this against Microsoft Excel,
;; replace progID to "Excel.Application".

(use win.ole)
(use gauche.collection)
(use gauche.array)

(define ko (make-ole "et.Application"))
(define book (? ko 'Workbooks ! 'Add))
(set! (~ ko 'Visible) #t)
(define sheet (? book 'Worksheets 1))
(set! (~ (? sheet 'range "A1") 'Value) "serial number")

(dynamic-wind (lambda() #f)
              (lambda()
                (let1 i 0
                  (for-each (lambda(x)(set! (~ x 'Value) i)(inc! i))
                            (? sheet 'Range "A2:E10")))
                (let1 a (? sheet 'Range "A2:E10" ? 'Value)
                  (array-for-each-index a
                    (^(i j) (format #t "array[~d,~d] = ~a~%"
                                    i
                                    j
                                    (array-ref a i j ))))))
              ole-release!)
