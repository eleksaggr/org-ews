;;; org-ews -- Synchronize calendar events from a Microsoft Exchange server
;;; -*- lexical-binding:t -*-

;;; Commentary:

;;; Code:
(require 'dom)
(require 'org-element)
(require 'parse-time)

(defcustom org-ews-host nil
  "The URI of the used Exchange server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-exchange-path "ews/exchange.asmx"
  "The path to the EWS API relative to `org-ews-host'."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-user nil
  "The username that is used to login to the Exchange server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-password nil
  "The password that is used to login to the Exchange server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-auth-mechanism
  '(:any :basic :plain :ntlm :negotiate)
  "The authentication mechanism used to login to the Exchange server."
  :type (list 'symbol)
  :group 'org-ews)

(defcustom org-ews-file nil
  "The file the calendar entries are stored in."
  :type 'string
  :group 'org-ews)

(defcustom org-ews--buffer-name "*EWS Sync*"
  "The name for the buffer that is used as output for `org-ews-sync', when `org-ews-file' is nil."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-insecure-connection nil
  "Whether the request to the Exchange server should be secured."
  :type 'boolean
  :group 'org-ews)

(defcustom org-ews-max-entries 100
  "The maximum amount of entries fetched in a request."
  :type 'integer
  :group 'org-ews)

(defcustom org-ews-days-past 7
  "The amount of days in the past to fetch entries for."
  :type 'integer
  :group 'org-ews)

(defcustom org-ews-days-future 30
  "The amount of days in the future to fetch entries for."
  :type 'integer
  :group 'org-ews)

(defcustom org-ews-sync-interval 300
  "The amount of seconds until `org-ews-sync' is called. \
If this is nil `org-ews-sync' is never called automatically."
  :type 'integer
  :group 'org-ews)

(defvar org-ews--timer nil
  "A timer that calls `org-ews-sync' upon expiry.")

(defcustom org-ews-curl-executable "curl"
  "The path to the cURL executable that is used to communicate with the Exchange server."
  :type 'string
  :group 'org-ews)

(defvar org-ews--curl-verbose nil
  "Whether cURL should be called with it's verbose flag set.")

(defvar org-ews--curl-charset "utf-8"
  "The charset of the cURL request.")


(defcustom org-ews--request-template "<?xml version=\"1.0\" encoding=\"utf-8\"?>
<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"
               xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"
               xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"
               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">
  <soap:Body>
    <FindItem Traversal=\"Shallow\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI=\"calendar:MyResponseType\"/>
        </t:AdditionalProperties>
      </ItemShape>
      <CalendarView MaxEntriesReturned=\"%d\" StartDate=\"%sT00:00:00-08:00\" EndDate=\"%sT00:00:00-08:00\"/>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id=\"calendar\"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>" "A template for a SOAP request to an Exchange server."
:type 'string
:group 'org-ews)

(defun org-ews-start ()
  "Start a periodic synchronization of calendar events.
The interval of synchronization is specified by `org-ews-sync-interval'.
Should `org-ews-sync-interval' be nil the synchronization is done once only."
  (interactive)
  (setq org-ews--timer (run-at-time nil org-ews-sync-interval 'org-ews-sync)))

(defun org-ews-stop ()
  "Stop the periodic synchronization of calendar events."
  (cancel-timer org-ews--timer))

(defun org-ews-sync ()
  "Fetch the upstream calendar events and save them.
If `org-ews-file' is nil, the formatted events are stored in the buffer
with the name `org-ews--buffer-name'.
Otherwise they are written to `org-ews-file'."
  (interactive)
  (let ((content (string-join (mapcar 'org-ews--format (mapcar 'org-ews--parse-fields (org-ews--parse-response (org-ews--execute-curl-request)))) "\n\n")))
    (if org-ews-file
        (with-temp-buffer
          (insert content)
          (write-region (point-min) (point-max) org-ews-file nil))
      (with-current-buffer (get-buffer-create org-ews--buffer-name) (insert content)))))

(defun org-ews--build-request ()
  "Build a SOAP request with the configured parameters.

In particular `org-ews-days-past' and `org-ews-days-future' are used to define
the range of calendar events that are fetched, and `org-ews-max-entries' is
used to limit the maximum amount of entries fetched."
  (let* ((get-date-from-delta (lambda (delta)
                                (org-read-date nil nil (format "+%d" delta))))
         (start-date (funcall get-date-from-delta (- org-ews-days-past)))
         (end-date (funcall get-date-from-delta org-ews-days-future)))
    (format org-ews--request-template org-ews-max-entries start-date end-date)))

(defun org-ews--build-target-url ()
  "Build the URL of the targeted Exchange server."
  (concat (if org-ews-insecure-connection "http" "https") "://" org-ews-host "/" org-ews-exchange-path))

(defun org-ews--build-credential-string ()
  "Builds the credential string for cURL."
  (concat org-ews-user ":" org-ews-password))

(defun org-ews--make-temp-request-file ()
  "Create a temporary file to store the formatted SOAP request in."
  (let ((temp-file (make-temp-file "org-ews")))
    (with-temp-buffer
      (progn
        (insert (org-ews--build-request))
        (write-region nil nil temp-file nil 1)
        temp-file))))

(defun org-ews--build-curl-command-string (request-file)
                                        ; TODO: This is really bad. Somebody that can see running processes will see the user's password.
  (let ((auth-param (cond ((eq org-ews-auth-mechanism :any) "--any")
                          ((eq org-ews-auth-mechanism :ntlm) "--ntlm")
                          (t "--any"))))
    (string-join (list
                  org-ews-curl-executable
                  (when org-ews--curl-verbose "-v")
                  "-s"
                  "-u" (org-ews--build-credential-string)
                  "-L" (org-ews--build-target-url)
                  "-H" (concat "\"" "Content-Type:text/xml; charset\\=" org-ews--curl-charset "\"")
                  "-d" (concat "@" request-file)
                  auth-param)
                 " "))
  )

(defun org-ews--execute-curl-request ()
  "Execute a cURL request built by `org-ews--build-curl-command-string'."
  (shell-command-to-string (org-ews--build-curl-command-string (org-ews--make-temp-request-file))))

(defun org-ews--parse-response (response)
  (with-temp-buffer
    (progn
      (insert response)
      (let ((content (libxml-parse-xml-region 1 (point-max)))) (dom-by-tag content 'CalendarItem)))))


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

(defun org-ews--parse-fields (dom)
  "Parse DOM and return all contained fields of an entry as an alist."
  (let ((entry (list)))
    (dolist (field org-ews--fields entry)
      (map-put entry (car field) (org-ews--parse-field dom (car field))))))

(defun org-ews--parse-field (dom field)
  "Parse DOM and return the value associated with FIELD.
FIELD should be a key from `org-ews--fields'."
  (let ((value (funcall (org-ews--get-field-parser field)
                        (dom-by-tag dom (eval (org-ews--get-field-tag field))))))
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
  (format-time-string "%FT%T%z" (parse-iso8601-time-string (dom-text dom))))

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
          (t nil))))

(defun org-ews--parse-organizer (dom)
  "Parse DOM and return its children as an assocation list representing its organizer property."
  (ignore))

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

(provide 'org-ews)
;;; org-ews.el ends here
