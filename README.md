# Generate Excel task list and gantt chart from redmine issues

## Japanese localisation version of redmine

The word "issue" is translated into "ticket" in the Japanese version of redmine. In this document, "issus" is used.

## Description

This is a Python script that extracts Redmine issues based on specified conditions and generates an Excel file from Redmine issues.  
The following Redmine information is used to generate the Excel task list and gantt chart.

- issue_id
- subject
- assigned_to
- start_date
- due_date
- closed_on
- done_ratio

If there are many issues that meet the specified conditions, processing will take a long time.

## Pre-requisites

Web service by REST API of Redmine must be enabled. Only the administrator of Redmine can enable it.  
See the section [Authentication](https://www.redmine.org/projects/redmine/wiki/Rest_api#Authentication) in the official site for details.

## How to use

1. Modify the 'config.toml' file refering the sample TOML file and to set the parameters for the gantt chart that you want to genarate.  
   Source data of redmine are also specified in the 'config.toml' file.
1. Run the Python script.
1. The script will generate a gantt chart in Excel format based on the redmine data.
1. After generating the gantt chart, the script asks you to input the excel base file name to save.  
   The extention of the file name is used as '.xlsx' automatically.

## Description of config.toml

Prepare 'config.toml' by referring the sample TOML file. The configuration file name must be 'config.toml'.

```
# config.toml

redmine.url = redmine server URL : ex. "https://redmine.org/"
redmine.project_name = project name : ex. "Project Blue"

redmine.account.need_login = set true if needs to login to redmine (refer below section for details)
redmine.account.username   = username for redmine account (refer below section for details)
redmine.account.password   = password for redmine account (refer below section for details)

redmine.filter.sort       = Column to sort. Append :desc to invert the order
redmine.filter.issue_id   = Find issue or issues by id (separated by ,)
redmine.filter.query_id   = Get issues for the given query id (refer below section for details)
redmine.filter.parent_id  = Get issues whose parent issue is given id
redmine.filter.tracker_id = Get issues from the tracker with given id
redmine.filter.status_id  = Get issues with given status id 
                            Possible values:
                              - "open" for open issues
                              - "closed" for closed issues
                              - "*" for all issues
                              - id for specific status id (ex. 1, 2, etc.)
redmine.filter.author_id  = Get issues which are authored by the given user id
redmine.filter.assigned_to_id   = Get issues which are assigned to the given user id
redmine.filter.fixed_version_id = Get issues with given version id

spreadsheet.font_name = font name that you want to use to excel : ex. "Meiryo UI"
spreadsheet.tab_title = excel tab title string : ex. "Project Blue"
                        If 'tab_title' is not specified, 'project_name' is used instead.

spreadsheet.gantt.start_date = start date for gantt chart in format "YYYY/MM/DD"
spreadsheet.gantt.end_date   = end date for gantt chart in format "YYYY/MM/DD"

holidays = [
  list of holidays in format "YYYY/MM/DD", ...
]
```

## Mandatory items of config.toml

At least the following items must be set in 'config.toml'.

```
  redmine.url
  redmine.project_name
  spreadsheet.gannt.start_date
  spreadsheet.gannt.end_date
```

## redmine.filter.query_id

This is the exclusive condition. If set this with other filters, other filters will be ignored.  
If this is a private query, it needs to login to redmine because the private query is connected to the account.

## Redmine account

It depends on if you need to login to extract issues from Redmine or not.  
If you don't need to login, set `redmine.account.need_login` to `false`. In this case, `redmine.account.username` and `redmine.account.password` will be ignored even those are set.  
If you need to login, set `redmine.account.need_login` to `true` and fill in the username and password. If `redmine.account.username` and/or `redmine.account.password` are empty, the script will prompt you to input them.
