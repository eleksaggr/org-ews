;;; org-ews-parser --- Parser for org-ews

;;; Commentary:

;;; Code:

(defvar org-ews-default-subject "Untitled Event")
(defvar org-ews-default-location "Unspecified Location")

(defvar org-ews--fields
  `((:subject . '(:parser org-ews--parse-subject :tag 'Subject :required t :default ,org-ews-default-subject))
   (:start-time . '(:parser org-ews--parse-start-time :tag 'Start :required t ))
   (:end-time . '(:parser org-ews--parse-end-time :tag 'End :required t))
   (:location . '(:parser org-ews--parse-location :tag 'Location :required nil :default ,org-ews-default-location))
   (:legacy-free-busy-status . '(:parser org-ews--parse-legacy-free-busy-status :tag 'LegacyFreeBusyStatus :required nil))
   (:item-type . '(:parser org-ews--parse-item-type :tag 'CalendarItemType :required nil))
   (:organizer . '(:parser org-ews--parse-organizer :tag 'Organizer :required nil))))

(defun org-ews--get-field-required (field)
  "Return the `:required' property of FIELD.
Field must exist in `org-ews--fields', otherwise nil is returned."
  (plist-get (caddr (assq field org-ews--fields)) :required))

(defun org-ews--get-field-parser (field)
  "Return the `:parser' property of FIELD.
Field must exist in `org-ews--fields', otherwise nil is returned."
  (plist-get (caddr (assq field org-ews--fields)) :parser))

(defun org-ews--get-field-default (field)
  "Return the `:default' property of FIELD.
Field must exist in `org-ews--fields', otherwise nil is returned."
  (plist-get (caddr (assq field org-ews--fields)) :default))

(defun org-ews--parse-field (dom field)
  (ignore))

(defun org-ews--parse-response (response)
  "Parse RESPONSE into a list of 'CalendarItem."
  (with-temp-buffer
    (insert response)
    (let ((content (libxml-parse-xml-region (point-min) (point-max))))
      (dom-by-tag content 'CalendarItem))))

(defun org-ews--parse-calendar-item (item-node)
  "Parse ITEM-NODE into an assocation list."
  (ignore))

(defconst org-ews-subject-fallback "Untitled Event"
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

(defconst org-ews-status-fallback :free
  "The fallback status for an event that does not have a status or whose status is an invalid value.")

(defun org-ews--parse-legacy-free-busy-status (node)
  "Parse NODE and return its LegacyFreeBusyStatus property.
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

(defconst org-ews-location-fallback "Unspecified"
  "The fallback location value for an event that does not have a location specified.")

(defun org-ews--parse-location (node)
  "Parse NODE and return its location property."
  (org-ews--parse-simple-text-field (dom-by-tag node 'Location) org-ews-location-fallback))

(defconst org-ews-item-type-fallback :single
  "The fallback item type value for an entry that does not have a valid item type.")

(defun org-ews--parse-item-type (node)
  "Parse NODE and return its calendar item type property.
Possible return values are `:single', `:occurence', `:exception'
or `:recurring-master'.
The values correspond to their meaning in the EWS specification.
If the text value does not fit to any of the above,
the value specified by `org-ews-item-type-fallback' is returned."
  (let ((text (org-ews--parse-simple-text-field (dom-by-tag 'CalendarItemType) nil)))
    (cond ((string= text "Single") :single)
          ((string= text "Occurence") :occurence)
          ((string= text "Exception") :exception)
          ((string= text "RecurringMaster") :recurring-master)
          (t org-ews-item-type-fallback))))

(defconst org-ews-my-response-type-fallback :unknown
  "The fallback MyResponseType value for an event that does not have a valid MyResponseType field.")

(defun org-ews--parse-my-response-type (node)
  "Parse NODE and return its MyResponseType property."
  (let ((text (org-ews--parse-simple-text-field (dom-by-tag node 'MyResponseType) nil)))
    (cond ((string= text "Unknown") :unknown)
          ((string= text "Organizer") :organizer)
          ((string= text "Tentative") :tentative)
          ((string= text "Accept") :accept)
          ((string= text "Decline") :decline)
          ((string= text "NoResponseReceived") :no-response-received)
          (t org-ews-my-response-type-fallback))))

(defun org-ews--parse-routing-type (node)
  "Parse NODE and return its RoutingType property."
  (ignore))

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
