;;; ~/.doom.d/org-ews.el -*- lexical-binding: t; -*-

(defcustom org-ews-host nil
  "The URI to the used Ews server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-exchange-path "ews/exchange.asmx"
  "The path to the EWS API relative to `org-ews-host'."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-user nil
  "The username that is used to login to the Ews server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-password nil
  "The password that is used to login to the Ews server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-auth-mechanism
  '(:any :basic :plain :ntlm :negotiate)
  "The authentication mechanism used to login to the Ews server."
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

(defcustom org-ews--timer nil
  "A timer that calls `org-ews-sync' upon expiry."
  :type 'timer
  :group 'org-ews)

(defcustom org-ews--curl-executable "curl"
  "The path to the cURL executable that is used to communicate with the Exchange server."
  :type 'string
  :group 'org-ews)

(defcustom org-ews--curl-verbose nil
  "Whether cURL should be called with it's verbose flag set."
  :type 'boolean
  :group 'org-ews)

(defcustom org-ews--curl-charset "utf-8"
  "The charset of the cURL request."
  :type 'string
  :group 'org-ews)


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

(defcustom org-ews--active-timestamps t
  "Whether timestamps should be active or inactive.
Do not modify this directly, use `org-ews-toggle-active-timestamps' instead."
  :type 'boolean
  :group 'org-ews)

(defcustom org-ews--timestamp-begin-symbol "<"
  "The symbol used to designate the begining of a timestamp."
  :type 'string
  :group 'org-ews)

(defcustom org-ews--timestamp-end-symbol ">"
  "The symbol used to designate the end of a timestamp."
  :type 'string
  :group 'org-ews)

(defun org-ews-toggle-active-timestamps ()
  "Toggle between creating active and inactive timestamps."
  (interactive)
  (progn
    (if org-ews--active-timestamps (setq org-ews--timestamp-begin-symbol "["
                                         org-ews--timestamp-end-symbol "]")
      (setq org-ews--timestamp-begin-symbol "<"
            org-ews--timestamp-end-symbol ">"))
    (setq org-ews--active-timestamps (not org-ews--active-timestamps))))

(defun org-ews-start ()
  "Call `org-ews-sync' and start a timer that calls `org-ews-sync' every interval specified by `org-ews-sync-interval'."
  (interactive)
  (let ((timer (run-at-time nil org-ews-sync-interval 'org-ews-sync)))
    (setq org-ews--timer timer)))

(defun org-ews-stop ()
  "Stop the periodic sync."
  (interactive)
  (cancel-timer 'org-ews--timer))

(defun org-ews--format-request ()
  "Format a SOAP request with the configured parameters."
  (let* ((get-date-from-delta (lambda (delta)
                                "Return the date `delta' days from today in Org date format."
                                (org-read-date nil nil (format "%+d" delta))))
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
  "Create a temporary file that stores the formatted SOAP request."
  (let ((temp-file (make-temp-file "org-ews")))
    (with-temp-buffer
      (progn
        (insert (org-ews--format-request))
        (write-region nil nil temp-file nil 1)
        temp-file))))                   ;

(defun org-ews--build-curl-command-string (request-file)
  ; TODO: This is really bad. Somebody that can see running processes will see the user's password.
  (let ((auth-param (cond ((eq org-ews-auth-mechanism 'any) "--any")
                          ((eq org-ews-auth-mechanism 'ntlm) "--ntlm")
                          (t "--any"))))
    (string-join (list
                  org-ews--curl-executable
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

(defun org-ews--parse-calendar-item (item-node)
  (list
   (org-ews--parse-subject (dom-by-tag item-node 'Subject))
   (org-ews--parse-start (dom-by-tag item-node 'Start))
   (org-ews--parse-end (dom-by-tag item-node 'End))
   (org-ews--parse-location (dom-by-tag item-node 'Location))
   (org-ews--parse-legacy-free-busy-status (dom-by-tag item-node 'LegacyFreeBusyStatus))
   (org-ews--parse-my-response-type (dom-by-tag item-node 'MyResponseType))
   (org-ews--parse-calendar-item-type (dom-by-tag item-node 'CalendarItemType))
   (org-ews--parse-has-attachments (dom-by-tag item-node 'HasAttachments))
   (org-ews--parse-organizer (dom-by-tag item-node 'Organizer))))

(defun org-ews--parse-item-id ()
  (ignore))

(defun org-ews--parse-simple-item (node)
  (dom-text node))

(defalias 'org-ews--parse-subject 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-start 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-end 'org-ews--parse-start)
(defalias 'org-ews--parse-legacy-free-busy-status 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-location 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-calendar-item-type 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-my-response-type 'org-ews--parse-simple-item)

(defun org-ews--parse-has-attachments (node)
  (string= (dom-text node) "true"))

(defun org-ews--parse-organizer (node)
  (list
   (org-ews--parse-name (dom-by-tag node 'Name))
   (org-ews--parse-email-address (dom-by-tag node 'EmailAddress))
   (org-ews--parse-routing-type (dom-by-tag node 'RoutingType))))

(defalias 'org-ews--parse-name 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-email-address 'org-ews--parse-simple-item)
(defalias 'org-ews--parse-routing-type 'org-ews--parse-simple-item)

(defcustom org-ews-attachment-tag "ATTACHMENT"
  "The tag that is placed on the entry when it has an attachment."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-fields '((:location . t)
                            (:response . t)
                            (:status . t)
                            (:has-attachment . nil)
                            (:event-type . nil)
                            (:organizer-name . t)
                            (:organizer-email . t)
                            (:organizer-routing-type . nil))
  "A list of fields to show for a calendar entry.
Each elements has the form (FIELD . INCLUDE), where FIELD is one of the possible fields
and INCLUDE is a boolean flag indicating whether to include the field when creating the
entry.")

(defcustom org-ews--entry-template "* %s
%s--%s
:PROPERTIES:
%s
:END:"
  "A template for a calendar entry in org-mode format."
  :type 'string
  :group 'org-ews)

(defcustom org-ews-untitled-event-placeholder "Untitled Event"
  "The subject of an entry whose name is the empty string."
  :type 'string
  :group 'org-ews)

(defun org-ews--format-time (time)
  ; TODO: Actually handle timezones here.
  (concat org-ews--timestamp-begin-symbol (replace-regexp-in-string "Z" "" (replace-regexp-in-string "T" " " time)) org-ews--timestamp-end-symbol))

(defun org-ews--format-entry (entry)
  "Format ENTRY."
  (let ((subject (if (string= (nth 0 entry) "") "Untitled Event" (nth 0 entry)))
        (start (org-ews--format-time (nth 1 entry)))
        (end  (org-ews--format-time (nth 2 entry))))
    (let* ((format-property (lambda (key prop-name value) (if (and (cdr (assoc key org-ews-fields)) (not (string= value "")))
                                                              (format ":%s: %s\n" prop-name value)
                                                            "")))
           (properties-string (replace-regexp-in-string "\n$" "" (concat (funcall format-property :location "LOCATION" (nth 3 entry))
                                      (funcall format-property :response "RESPONSE" (nth 4 entry))
                                      (funcall format-property :status "STATUS" (nth 5 entry))
                                      (funcall format-property :event-type "TYPE" (nth 6 entry))))))
      (format org-ews--entry-template subject start end properties-string))))

(defun org-ews-sync ()
  "Fetch the upstream calendar events and save them.
If `org-ews-file' is nil, the formatted events are stored in the buffer with the name `org-ews--buffer-name'.
Otherwise they are written to `org-ews-file'."
  (interactive)
  (let ((content (string-join (mapcar 'org-ews--format-entry (mapcar 'org-ews--parse-calendar-item (org-ews--parse-response (org-ews--execute-curl-request)))) "\n\n")))
        (if org-ews-file
            (with-temp-buffer
              (insert content)
              (write-region (point-min) (point-max) org-ews-file nil))
            (with-current-buffer (get-buffer-create org-ews--buffer-name) (insert content)))))
