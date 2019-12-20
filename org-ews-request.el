;;; ~/Code/org-ews/org-ews-request.el -*- lexical-binding: t; -*-

;;; Code:
(defconst org-ews-request--header "<?xml version=\"1.0\" encoding=\"utf-8\"?>
<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"
               xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"
               xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">
<soap:Body>"
  "The header for any SOAP requests made to the server.")

(defconst org-ews-request--footer "</soap:Body>
</soap:Envelope>"
  "The footer for any SOAP requests made to the server.")

(defconst org-ews-request--find-item "<FindItem
    xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"
    xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\"
    Traversal=\"%s\">
    <ItemShape>
        <t:BaseShape>%s</t:BaseShape>
        <t:AdditionalProperties>
            %s
        </t:AdditionalProperties>
    </ItemShape>
    %s
    <ParentFolderIds>
        <t:DistinguishedFolderId Id=\"%s\"/>
    </ParentFolderIds>
</FindItem>"

  "A template for a FindItem operation.")

(defun org-ews-request--prepare-find-item (traversal base-shape folder-id &optional additional-props calendar-view)
  "Prepare a FindItem operation with the given details.
TRAVERSAL may be either `:shallow', `:soft-deleted' or `:associated'.
BASE-SHAPE may be one of `:id-only', `:default' or `:all-properties'.
FOLDER-ID must be the distinguished name of a folder.
CALENDAR-VIEW is an optional `string' that contains a CalendarView XML element.
ADDITIONAL-PROPS is an optional `list' of `string' that are added to the
request as additional fields to be fetched. They are treated as FieldURI XML elements."
  (let ((field-uris (string-join (mapcar (lambda (uri) (format "<t:FieldURI FieldURI=\"%s\"/>" uri)) additional-props) "\n            "))
        (traversal-string (cl-case traversal
                            (:shallow "Shallow")
                            (:soft-deleted "SoftDeleted")
                            (:associated "Associated")
                            (t nil)))
        (base-shape-string (cl-case base-shape
                             (:id-only "IdOnly")
                             (:default "Default")
                             (:all-properties "AllProperties")
                             (t nil))))
    (when (cl-every #'identity `(,traversal-string ,base-shape-string ,folder-id))
        (format org-ews-request--find-item traversal-string base-shape-string field-uris folder-id))))

(defun org-ews-request--decorate-request (request)
  "Decorate REQUEST by adding a header and footer.
The header is specified in `org-ews-request--header', while the footer
is specified in `org-ews-request--footer'."
  (string-join `(,org-ews-request--header ,request ,org-ews-request--footer) "\n"))

(defun org-ews-request--execute (request)
  "Send REQUEST to the server using cURL and return the body of the response."
  (let ((tmp-file (let ((tmp-file (make-temp-file "ews")))
                     (with-temp-buffer
                     (insert request)
                     (write-region nil nil tmp-file nil 1))))
         (auth-mechanism (cl-case org-ews-auth-mechanism
                           (':any "--anyauth")
                           (':basic "--basic")
                           (':digest "--digest")
                           (':ntlm "--ntlm")
                           (':negotiate "--negotiate")
                           (t "--anyauth"))))
    (with-temp-buffer
      (progn
       (call-process org-ews-curl-executable nil t nil
                    "-s"
                    "-n"
                    ; TODO: Replace hard-coded strings of host URL with settings.
                    (format "%s://%s/ews/exchange.asmx" "https" org-ews-host)
                    "-L"
                    (format "-H Content-Type:text/xml;charset=%s" org-ews--curl-charset)
                    (format "-d @%s" tmp-file)
                    auth-mechanism))
      (buffer-string))))

(provide 'org-ews-request)
;;; org-ews-request.el ends here
