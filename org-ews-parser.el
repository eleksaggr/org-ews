;;; org-ews-parser --- Parser for org-ews

;;; Commentary:

;;; Code:
(require 'dom)
(require 'parse-time)

(defconst org-ews-default-subject "Untitled Event"
  "The default subject for a calendar entry that does not have a valid one.")

(defconst org-ews-default-location "Unspecified Location"
  "The default location for a calendar entry that does not have a valid one.")

(defconst org-ews-default-legacy-free-busy-status :free
  "The default legacy free busy status for a calendar entry that does not have a valid one.")

(defconst org-ews-default-my-response-type :unknown
  "The default my response type for a calendar entry that does not have a valid one.")

(defvar org-ews--fields
  `((:subject . '(:parser org-ews--parse-text-field :tag 'Subject :required t :default ,org-ews-default-subject))
    (:start-time . '(:parser org-ews--parse-time-field :tag 'Start :required t))
    (:end-time . '(:parser org-ews--parse-time-field :tag 'End :required t))
    (:location . '(:parser org-ews--parse-text-field :tag 'Location :required nil :default ,org-ews-default-location))
    (:legacy-free-busy-status . '(:parser org-ews--parse-legacy-free-busy-status :tag 'LegacyFreeBusyStatus :required nil :default ,org-ews-default-legacy-free-busy-status))
    (:my-response-type . '(:parser org-ews--parse-my-response-type :tag 'MyResponseType :required nil :default ,org-ews-default-my-response-type))
    (:item-type . '(:parser org-ews--parse-calendar-item-type :tag 'CalendarItemType :required nil))
    (:organizer . '(:parser org-ews--parse-organizer :tag 'Organizer :required nil))))

(defmacro org-ews--get-field-property (prop)
  "Define a function named `(concat \"org-ews--get-field-\" PROP)'.
The function returns the respective property, when given a field from `org-ews--fields'."
  `(defun ,(intern (concat "org-ews--get-field-" prop)) (field)
     ,(concat "Return the `:" prop "' property of FIELD.
Field must exist in `org-ews--fields', otherwise nil is returned.")
     (plist-get (caddr (assq field org-ews--fields)) ,(intern (concat ":" prop)))))

(org-ews--get-field-property "parser")
(org-ews--get-field-property "tag")
(org-ews--get-field-property "required")
(org-ews--get-field-property "default")
(org-ews--get-field-property "empty")

(defun org-ews--parse-field (dom field)
  "Parse DOM and return the value associated with FIELD.
FIELD should be a key from `org-ews--fields'."
  (let ((value (funcall (org-ews--get-field-parser field)
                        (dom-by-tag dom (org-ews--get-field-tag field)))))
    (if (not (org-ews--get-field-empty field))
        (unless (and (stringp value) (string= "" value)) value) value)))

(defun org-ews--parse-text-field (dom)
  "Parse DOM and return its text value."
  (dom-text dom))

(defun org-ews--parse-boolean-field (dom)
  "Parse DOM and return its text value as a boolean value.
If the value is `true' or `True' t is returned, other values result in nil."
  (let ((text (dom-text dom)))
    (or (string= "true" text) (string= "True" text))))

(defun org-ews--parse-time-field (dom)
  "Parse DOM and return its text value as a time string in `org-mode' format."
  (format-time-string "%Ft%T%z" (parse-iso8601-time-string (dom-text dom))))

(defun org-ews--parse-legacy-free-busy-status (dom)
  "Parse DOM and return its text value as a symbol representing the legacy free busy status.
Possible values for the legacy free busy status include:
- `:free'
- `:busy'
- `:no-data'
- `:out-of-office'
- `:tentative'
- `:working-elsewhere'
The meaning of these values corresponds to their definition in the EWS documentation."
  (let ((text (dom-text dom)))
    (cond ((string= "Free" text) :free)
          ((string= "Busy" text) :busy)
          ((string= "NoData" text) :no-data)
          ((string= "OOF" text) :out-of-office)
          ((string= "Tentative" text) :tentative)
          ((string= "WorkingElsewhere" text) :working-elsewhere)
          (t nil))))

(defun org-ews--parse-calendar-item-type (dom)
  "Parse DOM and return its text value as a symbol representing the calendar item type property.
Possible values for the calendar item type include:
- `:single'
- `:occurence'
- `:exception'
- `:recurring-master'
The meaning of these values corresponds to their meaning in the EWS documentation."
  (let ((text (dom-text dom)))
    (cond ((string= "Single" text) :single)
          ((string= "Occurence" text) :occurence)
          ((string= "Exception" text) :exception)
          ((string= "RecurringMaster" text) :recurring-master)
          (t nil))))

(defun org-ews--parse-my-response-type (dom)
  "Parse DOM and return its text value as a symbol representing the my response type property.
Possible values for the my response type property include:
- `:unknown'
- `:organizer'
- `:tentative'
- `:accept'
- `:decline'
- `:no-response-received'
The meaning of these values corresponds to their meaning in the EWS documentation."
  (let ((text (dom-text dom)))
    (cond ((string= "Unknown" text) :unknown)
          ((string= "Organizer" text) :organizer)
          ((string= "Tentative" text) :tentative)
          ((string= "Accept" text) :accept)
          ((string= "Decline" text) :decline)
          ((string= "NoResponseReceived" text) :no-response-received)
          (t org-ews-default-my-response-type))))

(defun org-ews--parse-organizer (dom)
  "Parse DOM and return its children as an assocation list representing its organizer property.
")

(provide 'org-ews-parser)
;;; org-ews-parser.el ends here
