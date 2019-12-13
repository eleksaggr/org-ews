;;; ~/Code/org-ews/org-ews-format.el -*- lexical-binding: t; -*-

;;; Commentary:

;;; Code:

(defconst org-ews--format-template
  "* %subject
:PROPERTIES:
%properties
:END:"
  "The template for a calendar entry in `org-mode' syntax.")

(defun org-ews--format-entry (entry)
  "Format ENTRY into an `org-mode' task."
  (format
   "* %s
%s--%s
:PROPERTIES:
%s
:END:")
  )

(defconst org-ews--format-property-list
  '((:location . "LOCATION")
    (:legacy-free-busy-status . "STATUS")
    (:my-response-type . "RESPONSE")
    (:item-type . "TYPE"))
  "A alist of properties that are included in the `:PROPERTIES' drawer.
The alist has the form (SYMBOL . PROPNAME),
where SYMBOL is the symbol to include,
and PROPNAME is the name of the respective property in the drawer.")

(defun org-ews--format-properties (entry)
  (string-join
   (remove-if 'null (mapcar (lambda (prop)
                         (org-ews--format-property prop (cdr (assq (car prop) entry))))
                       org-ews--format-property-list)) "\n"))

(defun org-ews--format-property (prop value)
  "Format PROP into an `org-mode' property with VALUE as value."
  (let ((pname (cdr prop)))
    (when value (concat ":" pname ":" " " value))))



(provide 'org-ews-format)
;;; org-ews-format.el ends here
