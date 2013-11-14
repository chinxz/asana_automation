#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import json
import subprocess
from datetime import date, datetime
import codecs
import xlsxwriter # https://xlsxwriter.readthedocs.org/en/latest/index.html

#PROJECT_ID='999999999999999'
#API_KEY='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
ASANA_URL='https://app.asana.com/api/1.0'
CSV_DELIM=','
CSV_QUOTE='"'
CSV_QESCP='""'
STORY_COMMENT='comment'
STORY_SYSTEM='system'
DEBUG = False

# Define comment class
class TaskStory:
	'The TaskStory class encapsulates the stories (comments, etc.) within Asana task'
	id = 0
	createdAt = None
	createdBy = ''
	type = ''
	text = ''
	
	def __init__(self, story_data):
		self.id = story_data['id']
		self.type = story_data['type']
		self.text = story_data['text']
		self.createdAt = parseDate(story_data['created_at'])
		person = story_data['created_by']
		if person != None:
			self.createdBy = person['name']

	def __str__(self):
		text = '[' + str(self.createdAt) + '] ' + self.createdBy + ': ' + self.text
		return text

	def __unicode__(self):
		text = u'[' + unicode(self.createdAt) + u'] ' + self.createdBy + u': ' + self.text
		return text
			
# Define a Task class
class AsanaTask:
	'The AsanaTask class encapsulates a task within Asana'
	id = 0
	name = ''
	parent = None
	createdAt = None
	dueOn = None
	modifiedAt = None  # modified_at
	description = '' # notes
	completed = False
	completedOn = None
	assignee = ''
	assigneeStatus = ''
	stories = None
	subTasks = None
	
	def __init__(self, task_json):
		task_data = json.loads(task_json)['data']
		self.id = task_data['id']
		self.name = task_data['name']
		# parent?
		self.createdAt = parseDate(task_data['created_at'])
		self.dueOn = parseDate(task_data['due_on'])
		self.modifiedAt = parseDate(task_data['modified_at'])
		self.description = task_data['notes']
		self.completed = task_data['completed']
		self.completedOn = parseDate(task_data['completed_at'])
		person = task_data['assignee']
		if person != None:
			self.assignee = person['name']
		self.assigneeStatus = task_data['assignee_status']
		
		# get stories
		self.getTaskStories()
		# get subtasks
		self.getSubTasks()

	def toRows(self, level=0):
		lines = []
		status = u''
		if self.completed: 
			status = u'Completed' 
		elif self.dueOn != None:
			# all timestamp in ISO format
			#duedate = datetime.strptime(self.dueOn, '%Y-%m-%d').date()
			if self.dueOn < date.today(): status = u'Overdue'
		
		task_name = self.name
		if level>0: task_name = u'    ' + task_name # indent sub-tasks, one level only; need to figure out a better way

		comments = None
		if self.stories.has_key(STORY_COMMENT):
			for comment in self.stories[STORY_COMMENT]:
				if comments==None:
					comments = unicode(comment)
				else:
					comments = comments + u'\n' + unicode(comment)

		line = (task_name, self.createdAt, self.assignee, self.dueOn, self.description, status, self.completedOn, comments)
		lines.append(line)

		if self.subTasks != None:
			for subtask in self.subTasks:
				lines.extend(subtask.toRows(level+1))
		return lines
		
	def getTaskStories(self):
		self.stories = {}
		# TODO Need to filter out system stories and only return comments
		task_story_cmd = ['curl', '-u', API_KEY + ': ', ASANA_URL + '/tasks/' + str(self.id) + '/stories']
		for line in run_command(task_story_cmd):
			pass
		if line != None:
			story_list = json.loads(line)['data']
			if story_list != None:
				for story_dict in story_list:
					story = TaskStory(story_dict)
					self.stories.setdefault(story.type, [])
					self.stories[story.type].append(story)
					if DEBUG: print u'added ' + story.type + u' story: ' + story.text + u' by: ' + story.createdBy
		if DEBUG: 
			if self.stories.has_key(STORY_COMMENT):
				print u'Totaly ' + str(len(self.stories[STORY_COMMENT])) + u' comments<<<'

	def getSubTasks(self):
		self.subTasks = []
		task_sub_cmd = ['curl', '-u', API_KEY + ': ', ASANA_URL + '/tasks/' + str(self.id) + '/subtasks']
		for line in run_command(task_sub_cmd):
			pass
		if line != None:
			subtask_list = json.loads(line)['data']
			if subtask_list != None:
				for subtask_dict in subtask_list:
					subtask = process_task(subtask_dict)
					if subtask != None: 
						self.subTasks.append(subtask)
						if DEBUG: print u'added subtask: ' + subtask.name + u' for: ' + subtask.assignee
					
					
# Running shell command using subprocess
def run_command(command):
	if DEBUG: print command
	p = subprocess.Popen(command,
						stdout=subprocess.PIPE,
						stderr=subprocess.STDOUT)
	return iter(p.stdout.readline, b'')

# Process a single task by fetching details/comments and all child tasks
def process_task(task_dict):
	task_id = task_dict['id']
	#task_name = task_dict['name']
	task_detail_cmd = ['curl', '-u', API_KEY + ': ', ASANA_URL + '/tasks/' + str(task_id)]

	for line in run_command(task_detail_cmd):
		pass
	
	if (line != None):
		task_detail = AsanaTask(line)
		return task_detail
	else:
		return None
	
def getHeader():
	line = (u'Task Name', u'Created On', u'Assignee', u'Due Date', u'Description', u'Status', u'Complete Date', u'Comments')
	return line

def joinCSVRow(line):
#	parts = ''
#	for part in line:
#		if part==None: part=''
#		parts = parts + part + CSV_DELIM
#	print parts
# 	is there a foreach function on list items?
	escapedRow = []
	for e in line:
		escapedRow.append(escapeCSVCell(e))
	return CSV_DELIM.join(escapedRow)

def escapeCSVCell(cell):
	if cell==None: return u''
	
	if not type(cell) is str: cell = unicode(cell)
	
	if cell.find(CSV_QUOTE)>0:
		cell = cell.replace(CSV_QUOTE, CSV_QESCP)
	
	# handle multi-line - NOTE IS THIS cross-platform?
	if cell.find(u'\r')>0:
		cell = cell.replace(u'\r', u'')
	
	if cell.startswith('-'):
		cell = "'" + cell

	cell = CSV_QUOTE + cell + CSV_QUOTE

	# other conditions???
	
	return cell

def parseDate(dt):
	if dt==None: return None
	
	dstr = None
	if dt.find('T'):
		dt = dt.split('T', 1)
		dstr = dt[0]
	else:
		dstr = dt
	
	return datetime.strptime(dstr, '%Y-%m-%d').date()

	
def writeWorkSheet(workbook, worksheet, task_rows):
	worksheet.set_zoom(90)
	# set line wrap on wide columns
	format_default = workbook.add_format()
	format_default.set_align('top')
	format_date = workbook.add_format()
	format_date.set_align('top')
	format_date.set_num_format('mm/dd/yyyy')
	format_multiline = workbook.add_format()
	format_multiline.set_align('top')
	format_multiline.set_text_wrap()
	# set column width
	worksheet.set_column('A:A', 50, format_multiline)
	worksheet.set_column('B:B', 11, format_date)
	worksheet.set_column('C:C', 15, format_default)
	worksheet.set_column('D:D', 11, format_date)
	worksheet.set_column('E:E', 50, format_multiline)
	worksheet.set_column('F:F', 11, format_default)
	worksheet.set_column('G:G', 11, format_date)
	worksheet.set_column('H:H', 50, format_multiline)
	# set up a table
	table_style={'style': 'Table Style Medium 2'}
	worksheet.add_table(0, 0, len(task_rows), 7, table_style) # 8 is the number of columns
	#add rows
	coln = 0
	rown = 0
	worksheet.write_row(0, 0, getHeader())
#	for h in getHeader():
		#worksheet.write(rown, coln, h)
		#coln += 1
	for row in task_rows:
#		coln = 0
		rown += 1
		worksheet.write_row(rown, 0, row)
		#		for h in row:
#			worksheet.write(rown, coln, h)
#			coln += 1

# convert an array of AsanaTask objects into array of text rows, each row is a tuple
def taskObjectsToTuples(task_objs):
	task_rows = []
	for task_obj in task_objs:
		# we convert the task to individual rows
		task_rows.extend(task_obj.toRows())
	return task_rows

###############################################################################
# parse command line arguments
# expects argv[1] = API_KEY
# argv[2] = PROJECT_ID
# argv[3] = Filename
if len(sys.argv) <= 3:
	sys.exit("Expect: getTask.py <API_KEY> <PROJECT_ID> <OUTPUT_FILE>")

API_KEY = sys.argv[1]
PROJECT_ID = sys.argv[2]
filename = sys.argv[3].lower()

# get the tasks
task_cmd = ['curl', '-u', API_KEY + ': ', ASANA_URL + '/projects/' + PROJECT_ID + '/tasks']
#print task_cmd

for line in run_command(task_cmd):
	pass
task_data = json.loads(line)
#print(task_data)

# iterate JSON dict object to obtain the task list
# {u'data': [{u'id': 8231939798046, u'name': u'RCA outage on DB 10/17'}, 
#        {u'id': 8216164833899, u'name': u'jssl_497_This_That'}, ...]}
task_list = task_data['data']
open_objs = []
closed_objs = []
for task in task_list:
	# task_obj is an Asana_Task object
	task_obj = process_task(task)
	if task_obj.completed:
		closed_objs.append(task_obj)
	else:
		open_objs.append(task_obj)

# sort open tasks by open date in descending order
open_objs = sorted(open_objs, key=lambda task: task.createdAt, reverse=True)   # sort by createdAt
# sort closed task by close date in descending order
closed_objs = sorted(closed_objs, key=lambda task: task.completedOn, reverse=True)   # sort by completedOn

# write to file
if filename.endswith('xlsx') or filename.endswith('xls'):
	workbook = xlsxwriter.Workbook(filename)
	worksheet_open = workbook.add_worksheet('Open Tasks ' + str(date.today()))
	writeWorkSheet(workbook, worksheet_open, taskObjectsToTuples(open_objs))
	worksheet_closed = workbook.add_worksheet('Closed Tasks ' + str(date.today()))
	writeWorkSheet(workbook, worksheet_closed, taskObjectsToTuples(closed_objs))
	# worksheet
	workbook.close()
else:
	# assume it is CSV
	fo = codecs.open(filename, 'wb', 'utf-8')

	# print header row
	fo.write(joinCSVRow(getHeader()))
	fo.write('\r\n')

	for row in taskObjectsToTuples(open_objs):
		fo.write(joinCSVRow(row))
		fo.write('\r\n')

	for row in taskObjectsToTuples(closed_objs):
		fo.write(joinCSVRow(row))
		fo.write('\r\n')

	fo.close()
