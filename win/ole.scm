;;;
;;; ole
;;;

(define-module win.ole
  (export make-ole ole-connect ole-call-method ole?
          ole-set! ole-ref ole-release! <ole> <ole-condition> <ole-collection>
          ole-condition? ole-condition-number ole-const ole-methods
          ! ?))
(select-module win.ole)

(use gauche.collection)
(use gauche.array)
(dynamic-load "ole")

(define-condition-type <ole-condition> <message-condition> ole-condition?
  (number ole-condition-number))

(define-method object-apply ((obj <ole>) (method <symbol>) . args)
  (apply ole-call-method obj method args))

(define-method ref ((obj <ole>) (prop <symbol>))
  (ole-ref obj prop))

(set! (setter ole-ref) (lambda(obj prop val) (ole-set! obj prop val)))

(define-method (setter ref) ((obj <ole>) (prop <symbol>) val)
  (ole-set! obj prop val))

(define-method call-with-iterator ((obj <ole-collection>) proc :allow-other-keys)
  (let* ((iter (make-ole-iterator obj))
         (buf (ole-iterator-next iter)))
    (let ((end? (lambda()(eof-object? buf)))
          (next (lambda()(rlet1 r buf (set! buf (ole-iterator-next iter))))))
      (proc end? next))
    (ole-iterator-release iter)))

;; method chain
(define-syntax !
  (syntax-rules ()
    ((_ x ...)
     (!-helper () x ...))))

(define-syntax !-helper
  (syntax-rules (! ?)
    ((_ (as ...))
     (as ...))
    ((_ (as ...) ! rest ...)
     (!-helper ((as ...)) rest ...))
    ((_ (as ...) ? rest ...)
     (!-helper (ole-ref (as ...)) rest ...))
    ((_ (as ...) a rest ...)
     (!-helper (as ... a) rest ...))
    ))

(define-syntax ?
  (syntax-rules ()
    ((_ x ...)
     (?-helper (ole-ref) x ...))))

(define-syntax ?-helper
  (syntax-rules (? !)
    ((_ (as ...))
     (as ...))
    ((_ (as ...) ? rest ...)
     (?-helper (ole-ref (as ...)) rest ...))
    ((_ (as ...) ! rest ...)
     (?-helper ((as ...)) rest ...))
    ((_ (as ...) a rest ...)
     (?-helper (as ... a) rest ...))
    ))
