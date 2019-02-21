### Monitor your office 365 Environment for newly created Identities.
#### This report will fetch all the new users setup in the office 365 environment in Last 24 hours. and will share the analyses of the new user typed created.
##### This report will also alert on New Cloud only account to monitor any unauthrised account setup in the environment.

##### SAMPLE REPORT BELOW

|USER Type| Numbers |
|---|---|
|Type UserMailbox Created	|**48**|
|Type Mailuser Created	|**20**|
|Type GuestUser Created - (External users added to Teams or O365 Group)|	**48**|
|Type user Created	|**66**|
|Type Shared Created	|**7**|
|Total Number of Users Created	|**150**|

**Below Cloud Only Account were Found** 

|WhenCreatedUTC|	DisplayName|	RecipientTypeDetails	|IsDirSynced |
|---|---|---|---|
|20-02-2019 17:15:14| Admin-test |	User	|False |
|21-02-2019 09:43:59	|SVC-Test	|User|	False|
|21-02-2019 09:44:17	| test|	User|	False|

