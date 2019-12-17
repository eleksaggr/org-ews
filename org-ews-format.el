;;; org-ews-format --- Format calendar entries into `org-mode' entries

;;; Commentary:

;;; Code:
(require 'org-element)

(defconst org-ews--properties '((:item-type . "TYPE")
                                (:my-response-type . "RESPONSE")
                                (:legacy-free-busy-status . "STATUS")
                                (:location . "LOCATION")
                                )
  "An alist that maps fields of an entry to their names as `org-mode' property.
The alist has the form (FIELD . PNAME), where FIELD is a field
from `org-ews--fields' and PNAME is the name of the corresponding
property in a calendar entry, when formatted as an `org-mode' entry.")

(defvar org-ews--timestamp-begin-symbol "<"
  "The symbol used to designate the begining of a timestamp.")

(defvar org-ews--timestamp-end-symbol ">"
  "The symbol used to designate the end of a timestamp.")

(defvar org-ews-active-timestamps t
  "Whether timestamps should be active or inactive.
Do not modify this directly, use `org-ews-toggle-active-timestamps' instead.")

(defun org-ews-toggle-active-timestamps ()
  "Toggle between designating timestamps as active or inactive.
The current state can be seen by examining `org-ews-active-timestamps'.

Prefer using this function over modifying `org-ews-active-timestamps' directly."
  (interactive)
  (progn
    (if org-ews-active-timestamps (setq org-ews--timestamp-begin-symbol "["
                                        org-ews--timestamp-end-symbol   "]")
      (setq org-ews--timestamp-begin-symbol "<"
            org-ews--timestamp-end-symbol   ">"))
    (setq org-ews-active-timestamps (not org-ews-active-timestamps))))

(defun org-ews--format (entry &optional quiet)
  "Format ENTRY into an `org-mode' entry.
ENTRY should be an alist of the form (FIELD . VALUE),
where FIELD is a symbol identifying a field from `org-ews--fields'
and VALUE is the `string' value of the field.

If QUIET is non-nil a message is printed, when formatting fails."
  (let ((subject (alist-get :subject entry))
        (start-time (alist-get :start-time entry))
        (end-time (alist-get :end-time entry)))
    (if (and (and start-time end-time) subject)
                                        ; TODO: Don't take a fixed template here.
        (format "* %s\n%s--%s\n:PROPERTIES:\n%s\n:END:\n"
                subject
                (concat org-ews--timestamp-begin-symbol start-time org-ews--timestamp-end-symbol)
                (concat org-ews--timestamp-begin-symbol end-time org-ews--timestamp-end-symbol)
                (string-join (org-ews--format-properties entry) "\n"))
      (unless quiet (error "Entry is missing one of the required fields: :subject, :start-time or :end-time")))))

(defun org-ews--format-properties (entry)
  "Format fields in ENTRY into `org-mode' properties.
This returns a list of `string's that are of the form \":PROP: VALUE\",
where PROP is the name of the respective property in `org-ews--properties'
and VALUE is the value of the property in ENTRY."
  (let ((properties (list)))
    (dolist (field org-ews--properties properties)
      (let ((prop-name (cdr field))
            (prop-value (cdr (assq (car field) entry))))
        (when prop-value (push (concat ":" prop-name ":" " " prop-value) properties))))))

(provide 'org-ews-format)
;;; org-ews-format.el ends here
