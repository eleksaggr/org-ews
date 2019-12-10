;;; org-ews-parser --- Parser for org-ews

;;; Commentary:

;;; Code:

(defun org-ews--parse-response (response)
  "Parse RESPONSE into a list of 'CalendarItem."
  (with-temp-buffer
    (insert response)
    (let ((content (libxml-parse-xml-region (point-min) (point-max))))
      (dom-by-tag content 'CalendarItem))))

(defun org-ews--parse-calendar-item (item-node)
  "Parse ITEM-NODE into an assocation list."
  (ignore))

(defvar org-ews-subject-fallback "Untitled Event"
  "The fallback subject for an event that does not have a title or whose title is the empty string.")

(defun org-ews--parse-subject (node)
  "Parse NODE and return its subject."
  (org-ews--parse-simple-text-field (dom-by-tag node 'Name) org-ews-subject-fallback))

(defun org-ews--parse-start-time (node)
  "Parse NODE and return its starting time.
Since there is no sane fallback value for this field we return nil upon error instead."
  (org-ews--parse-simple-text-field (dom-by-tag node 'Start) nil))

(defun org-ews--parse-end-time (node)
  "Parse NODE and return its ending time.
Since there is no sane fallback value for this field we return nil upon error instead."
  (org-ews--parse-simple-text-field (dom-by-tag node 'End) nil))

(defvar org-ews-status-fallback :free
  "The fallback status for an event that does not have a status or whose status is an invalid value.")

(defun org-ews--parse-status (node)
  "Parse NODE and return its status property.
Possible return values are `:free', `:busy', `:no-data',
`:out-of-office', `:tentative', `:elsewhere'.
The values correspond to their meaning in the EWS specification.
If the text value does not fit to any of the above,
the value specified by `org-ews-status-fallback' is returned."
  (let ((text (org-ews--parse-simple-text-field (dom-by-tag node 'LegacyFreeBusyStatus) nil)))
    (cond ((string= text "Free") :free)
          ((string= text "Busy") :busy)
          ((string= text "NoData") :no-data)
          ((string= text "OOF") :out-of-office)
          ((string= text "Tentative") :tentative)
          ((string= text "WorkingElsewhere") :elsewhere)
          (t org-ews-status-fallback))))

(defun org-ews--parse-simple-text-field (node &optional fallback)
  "Parse NODE and return its text value.
FALLBACK is returned instead, if either there is no text value
or the text value is the empty string."
  (let ((text (dom-text node)))
        (if (or (not text) (string= "" text)) fallback text)))

(defun org-ews--parse-boolean-text-field (node)
  "Parse NODE and return its text value as a boolean value.
A text value of \"true\" is considered t, while other values are considered nil."
  (string= "true" (dom-text node)))

(provide 'org-ews-parser)
;;; org-ews-parser.el ends here
