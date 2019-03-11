### This script gets all the messages sent by a user on a specific date can also include stats for sent to domains.
#### The script is fully tested for Exchange Online to the script publising date.

### How to run Example:
``` Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5```

### Use Switch ``` "-GetSentToDomainStats" ``` to includes stats for email sent to domains
``` Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5 -GetSentToDomainStats ```

### Use Switch ``` "-ShowProgress" ``` will show progress while script is running.
```Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5 -GetSentToDomainStats -ShowProgress ```

### Use Switch ``` "-DeepAnalysis" ``` to include the message sent per minute, also include the top domain by default.
```Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5 -DeepAnalysis -ShowProgress ```
