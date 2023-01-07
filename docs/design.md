# Race Scoring High Level Design


The overall design aims to accomodate a simple scoring and results publishing system.  This will be accomplished by the following:
1.  Support importing raw race results via xlsx or csv
2.  Storing and maintaining supporting data in xlsx, csv or text format
3.  A scoring utility to process raw results and supporting data
4.  A presentation utilty (might be part of scoring utility) that generates web output (html, css, js) suitable for presenting by a website.
5.  Hierarchical file basedsStorage

##  1.  Support Importing Raw Results
The raw results will come in via 3rd party data either in xlsx or csv.  It must have the following standard set of columns: 

	["RaceColumn", "Place", "Sex", "FName", "LName", "Age", "City", "State", "Time", "USATFResult"]


The raw data may be manually massaged into this format for use by the scoring system.


## 2. Supporting Data
The following supporting data must be stored in xlsx, csv, or text format:

	A.  Membership information
	B.  Team rosters
	C.  Series Information (races in the series)
	D.  (?)Scoring tables


## 3. Scoring Utility
The scoring utility should be able to generate:

	A.  Individual race results.
	B.  Series results.

The scoring utility should have direct access to items described in #1 and #2.  It should also be capable of processing 1, 2A and 2B to generate race results (3A above) as well as series results (3B) from items 2C and 3A.


## 4.  Presentation Utility

The presentation utility should have access to the output of 3A and 3B and be able to generate output comparable to the html examples that can be found in [examples|./examples/].  The utility may be part of the scoring utility, or separate. 


## 5.  All of these files should be stored in a hierarchical file-like format.  It should support remote storage, backup an preferably versioning.  Options include:
 - Google Drive
 - Box
 - Github

### Github
- Would be great, but may require too much technical expertise.  
- A free account would be fine.  
- The versioning serves as an excellent backup.

### Google drive 
- is easy for sharing and accessing
- The free tier provides sufficient space + sharing, 
- Lacks good versioning.  
- Backup would be manual and somewhat error prone.

### Box 
 - has a good free account, 
 - but versioning is a bit weak
 - Backup is likely very manual.
