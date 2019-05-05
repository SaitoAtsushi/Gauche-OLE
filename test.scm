;;;
;;; Test win.ole
;;;

(use gauche.test)

(test-start "win.ole")
(use win.ole)
(test-module 'win.ole)

;; The following is a dummy test code.
;; Replace it for your tests.

;; (test* "test-ole" "ole is working"
;;        (test-ole))

;; If you don't want `gosh' to exit with nonzero status even if
;; the test fails, pass #f to :exit-on-failure.
(test-end :exit-on-failure #t)
