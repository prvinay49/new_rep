# BranchComparator

Branch comparator is an utility to check the missing changes from one branch to another.

It basically checks all the repos for the given source and destination branch and we can also make it run for specified devices.
We can also specify date range to check for the changes.

### Required External Packages

  - pygerrit2
  - xlwt
  - xmltodict
  
### Configuration

Directory for config files: config/

Files:
  - devices.json [Contains devices list with execution flag, required devices execution flags should be set if the device_specific_flag is set in the execution]
  - gerrit.json [Contains gerrit credentials]
  - manifest.json [Contains project manifest details for the devices]
  
Note: modify gerrit.json file with your own credentials before running the script and also ensure that you are connected to VPN
### Usage

Usage: python branch_comparator.py <source_branch> <destination_branch> <device_specific_flag> [<start_time> <end_time>]

<start time> and <end time> should be given in IST time zone
  
  Date time range parameters are optional

Example with Date time range parameters:
```sh
$ python branch_comparator.py 3.10_p1v stable2 false 2019-07-10-00:00:00 2019-07-31-11:00:00
```

Example without Date time range parameters:
```sh
$ python branch_comparator.py 3.10_p1v stable2 false
```

Example with device specific flag enabled:
```sh
$ python branch_comparator.py 3.10_p1v stable2 true 2019-07-10-00:00:00 2019-07-31-11:00:00
```
