import pprint
import sys

import win32com.client

# folder constants
olFolderTodo = 4
olFolderTasks = 13

# item constants
olTaskItem = 48

task_fields = ['ActualWork',
               'Application',
               'AutoResolvedWinner',
               'BillingInformation',
               'Body',
               'CardData',
               'Categories',
               'Class',
               'Companies',
               'Complete',
               'ContactNames',
               'Contacts',
               'ConversationID',
               'ConversationIndex',
               'ConversationTopic',
               'CreationTime',
               'DateCompleted',
               'DelegationState',
               'Delegator',
               'DownloadState',
               'DueDate',
               'EntryID',
               'FormDescription',
               'Importance',
               'InternetCodepage',
               'IsConflict',
               'IsRecurring',
               'LastModificationTime',
               'Links',
               'MarkForDownload',
               'MessageClass',
               'Mileage',
               'NoAging',
               'Ordinal',
               'OutlookInternalVersion',
               'OutlookVersion',
               'Owner',
               'Ownership',
               'Parent',
               'PercentComplete',
               'ReminderOverrideDefault',
               'ReminderPlaySound',
               'ReminderSet',
               'ReminderSoundFile',
               'ReminderTime',
               'ResponseState',
               'Role',
               'Saved',
               'SchedulePlusPriority',
               'SendUsingAccount',
               'Sensitivity',
               'Session',
               'Size',
               'StartDate',
               'Status',
               'StatusOnCompletionRecipients',
               'StatusUpdateRecipients',
               'Subject',
               'TeamTask',
               'ToDoTaskOrdinal',
               'TotalWork',
               'UnRead'
               ]

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

# get all the tasks
tasks_folder = ns.GetDefaultFolder(olFolderTasks)
task_items = tasks_folder.Items
print("Number of tasks:", len(task_items))

# find the names of all user defined fields
user_field_names = []
user_defined_fields_collection  = tasks_folder.UserDefinedProperties
for i in range(user_defined_fields_collection.Count):
    user_field = user_defined_fields_collection[i].Name
    user_field_names.append(user_field)
    print("user field:", user_field)


def dump_task(task_item):
    row = {}
    # get the "regular" fields
    for f in  task_fields:
        row[f] = getattr(task_item, f)
    # get any user defined fields
    properties = task_item.UserProperties
    for i in range(properties.Count):
        row[properties[i].Name] = repr(properties[i])
    return row


pprint.pprint(dump_task(task_items[0]))


for i in range(0, len(task_items)):
    item = task_items[i]
    print(item.Subject)


