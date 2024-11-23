import openpyxl
import pylightxl as xl
import calendar
import datetime
import string

MAX = 5
week_total = 4
# Semanal, Bisemanal, Trisemanal

class Task:
     def __init__(self, id, description, frequency, duration, place):
        self.id = id
        self.description = description
        self.frequency = frequency
        self.duration = duration
        self.place =  place
        self.counter=0# count the number of times task has been executed

class Week:
     def __init__(self,  week_number, tasks_duration):
        self.week_number = week_number
        self.tasks_duration = tasks_duration
        self.tasks = [] # list of task ids

def register_planned_task(task_id):
  if task_id in plannedTasks:
    return
  else:
   plannedTasks[task_id]=1

def count_planned_tasks(): 
  print("#planned tasks is", len(plannedTasks), "and #total tasks is ", len(tasks))

def storePlace(place):
    if not any(place in x for x in placesList):
       placesList.append(place)
      
def storeTasks():
    i=0
    list_tasks_id_per_freq = []      
    # each excel line is a list containing task's description, frequency, duration and place
    for x in db.ws('TarefasPY').rows: 
      if  not str(x[0]) or  not str(x[1])  or  not str(x[0]) or not str(x[3]): # discard incomplete lines and header line     
        continue
      id = i
      desc = x[0]
      frequency = x[1]
      duration = x[3]
      place = x[2] 
      storePlace(place)
      task = Task(id, desc, frequency, duration, place)
      tasks[id] = task
      if str(x[1]) in freq: # if frequency already exists , add task id to frequency's tasks list
          list_tasks_id_per_freq = freq[frequency] # gets tasks ids list  
          list_tasks_id_per_freq.append(id)
          freq[frequency] = list_tasks_id_per_freq # aqui
      else:
          list_tasks_id_per_freq = [id]   
          freq[frequency]= list_tasks_id_per_freq
    #Debug: print ('Task #',task.id, len(str(x[0])) ,' name ', task.description, ' duration', task.duration)        
      i = i+1

def weekly_Tasks(frequency, week_tasks_duration):
    for task in freq[frequency]: 
       week_tasks_duration = week_tasks_duration + task.duration 
    print('   Total weekly tasks duration is ', week_tasks_duration)
    return week_tasks_duration

def frequency_period(frequency):
    match frequency:
        case 'Quinzenal':
            return 2
        case 'Mensal':
            return 5
        case 'Bimestral':
            return 9
        case 'Trimestral':
            return 14
        case 'Semestral':
            return 29
        case 'Anual':
            return 55
    
def orderTasksByPlace( ):
  # transforms  value of freq dictionary into a list of task objects ordered by place 

  for frequency in freq.keys(): # according to tasks frequency 
  # place can be Cozinha, Casa, Casa de Banho, Quartos, Varanda, Marquise #6 tipos 
   orderedTasksList =[]
   for task_id in freq[frequency]:
     task_place = tasks[task_id]
     orderedTasksList.append(task_place)
  
   new_list = sorted(orderedTasksList, key=lambda x: x.place, reverse=True)
   freq[frequency] = new_list
  #  print(frequency)
  #  for i in freq[frequency] :
  #    print("    ",i.place, ",",i.id )
            
  print("-------------------------------------")


def taskExistsInWeek(task_id, week_number):
  for t in tasksperWeek[week_number]:
    if task_id == t.id:
         return True
  return False    

def allocateTasks(curr_week_number): # one week's tasks according to frequency, duration, place in these order     
  week_tasks_duration = 0
  week_task_list = [] # current week tasks
  print('>>>>>>>>>>>>>>>>>>>>>>>>>> week ' ,curr_week_number)
  
  for frequency in freq.keys(): #allocate according to tasks frequency 
    if round(week_tasks_duration) == MAX:
       tasksperWeek[curr_week_number]= week_task_list.copy()
       break
    # print("week tasks", frequency, len(week_task_list))############
    if frequency == 'Semanal': 
      for task in freq[frequency]: 
        week_tasks_duration = week_tasks_duration + task.duration 
        register_planned_task(task.id)
        week_task_list.append(task)
    else: 
      # print("weel tasks list line120 ", len(week_task_list))############
      interval = frequency_period(frequency)      
       # interval defines the interval of prior weeks where the task should be executed
      for task in freq[frequency]: 
        done = 0 
        if curr_week_number > interval :  
          prior_week_interval =  curr_week_number - interval  
          # print("task", task.id,' debug: interval', interval, 'prior_week_accoridng _interval_', prior_week_interval)############
          if taskExistsInWeek(task.id, prior_week_interval): 
          #if task was performed precisely on the interval ,then it MUST be allocateon the current week
            week_task_list.append(task)
            week_tasks_duration = week_tasks_duration + task.duration
            register_planned_task(task.id)
            continue

          # #if task was performed previously to the interval , MUST be allocate TODO relate to interval th e condition
        prior_week_number = curr_week_number - 1
        while  prior_week_number > 0 and curr_week_number - prior_week_number <= interval:  
          if taskExistsInWeek(task.id, prior_week_number): 
             # task performed previously 
             done = 1  
             break
          prior_week_number = prior_week_number - 1                
                
        #if task was not performed on the interval, try allocate
        if done == 0: 
           week_tasks_duration = week_tasks_duration + task.duration            
           if week_tasks_duration  <= MAX: 
              # task is allocated 
              week_task_list.append(task)
              register_planned_task(task.id)
           else:
              # task is not allocated
              week_tasks_duration = week_tasks_duration - task.duration 
              if round(week_tasks_duration) == MAX: # if maximum weekly time for tasks has been reached then stop
                 tasksperWeek[curr_week_number]= week_task_list.copy()
                 break
        
  tasksperWeek[curr_week_number]= week_task_list.copy()
  print('>>>>>>>>>>>>>>>>>>>>>>>>>>',len(tasksperWeek[curr_week_number]),' tasks','| round tasks duration:',round(float(week_tasks_duration)),' hours') 
  for t in tasksperWeek[curr_week_number]:
    if t.frequency == 'Semanal':
      continue
    print("   ",t.frequency,t.place,t.id)
    # print(tasksperWeek[curr_week_number])
    # week = Week(week_number, week_tasks_duration, week_task_list)

def create_column_list():
   
    #list of 26 first elements list(string.ascii_uppercase)
    
    number_of_cols = len( tasksperWeek) 
    list_of_cols_names = []
    
    list_of_cols_names =list(string.ascii_uppercase)
  
    counter = len(list_of_cols_names)
    i = 0

    while counter < number_of_cols:## assuming number of weeks is >26
     for col in list(string.ascii_uppercase):
       #print(list_of_cols_names[i] + col)
       column=list_of_cols_names[i] + col
       list_of_cols_names.append(column)
       counter = counter +1
       if counter >= number_of_cols:
          break 
     i=i+1
  #  print( 'number of weeks is ', counter, 'and  column list is ', list_of_cols_names, 'size is', len(list_of_cols_names))
    return list_of_cols_names

   
         
def write_calendar_to_excel():
    wb = openpyxl.load_workbook('/Users/maf/Desktop/tarefas2.xlsx')
    mysheet = wb.create_sheet(index=4, title='St')
    
    cols = create_column_list() 
    # week number , tasks 
    for week in sorted(tasksperWeek.keys(),reverse = True):
        column_name = cols.pop()
        i=1
        mysheet[ str(column_name) +str(i)] = 'week ' + str(week)
     #   mysheet[cell].value
        i = i + 1
        for task in tasksperWeek[week]:
          print(column_name , 'and ', i)
          mysheet[str(column_name) +str(i)] =  str(task.place) +': '+ str(task.description)  
          i = i + 1
    wb.save('/Users/maf/Desktop/tarefas2.xlsx')

################### ################### ################### 
###################         MAIN        ################### 
################### ################### ################### 

db = xl.readxl(fn='/Users/maf/Desktop/tarefas.xlsx')

freq = dict() # key: task's execution frequency, value: list of task_ids
tasks = dict()# key:task_id, value: Object Task

tasksperWeek = dict() ##  key: week_number, value: Object Week list of tasks_ids and
plannedTasks = dict() ## key: task_id , value: number of task occurences planned 
placesList = [] # list of places to  

storeTasks() # fill freq and tasks dictionaries 
print(  '------------------------- Total tasks number is' , len(tasks))

orderTasksByPlace()
tasks.clear() # dictionary is no longer needed

i=1
while i <= week_total:
  allocateTasks(i)
  i = i+1

# ## Allocate tasks to weeks concerning conditions:
# ### frequency is repected
# ### tasks duration <= MAX hours  

# count_planned_tasks()
# #print("planned", sorted(plannedTasks.keys()))

#write_calendar_to_excel()

# current number in one year
current_week = datetime.date(datetime.date.today().year, datetime.date.today().month, datetime.date.today().day).strftime("%V")
#next_week= current_week_other + datetime.timedelta(days=7)

today = datetime.date.today()
# today = today - datetime.timedelta(days=2)
if today.weekday() > 0:
   monday = today - datetime.timedelta(days=today.weekday())
else:
   monday = today 
#print( 'today is', today, today.weekday(), 'week started at ',monday )
# next week monday
print('curr week is ', current_week,' started at monday', monday,'next monday is ', monday + datetime.timedelta(weeks=1))

