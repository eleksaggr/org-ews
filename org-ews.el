;;; org-ews -- Synchronize calendar events from a Microsoft Exchange server

;;; Commentary:

;;; Code:
(require 'org-ews-format)
(require 'org-ews-parse)

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
  (let ((content (string-join (mapcar 'org-ews--format (mapcar 'org-ews--parse-calendar-item (org-ews--parse-response (org-ews--execute-curl-request)))) "\n\n")))
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
        (insert (org-ews--format-request))
        (write-region nil nil temp-file nil 1)
        temp-file))))

(defun org-ews--build-curl-command-string (request-file)
                                        ; TODO: This is really bad. Somebody that can see running processes will see the user's password.
  (let ((auth-param (cond ((eq org-ews-auth-mechanism 'any) "--any")
                          ((eq org-ews-auth-mechanism 'ntlm) "--ntlm")
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

;; (defun org-ews--parse-response (response)
;;   (with-temp-buffer
;;     (progn
;;       (insert response)
;;       (let ((content (libxml-parse-xml-region 1 (point-max)))) (dom-by-tag content 'CalendarItem)))))



(provide 'org-ews)
;;; org-ews.el ends here
