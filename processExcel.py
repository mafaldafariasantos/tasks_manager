import openpyxl
import pylightxl as xl
import calendar
import datetime

MAX = 8
week_total=3
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

def storeTasks():
    i=0
    task_taskdur = [] 
    # each excel line is a list containing task's description, frequency, duration and place
    for x in db.ws('TarefasPY').rows: 
      if  not str(x[0]) or  not str(x[1])  or  not str(x[0]) or not str(x[3]): # discard incomplete lines and header line     
        continue
      
      task = Task(i, x[0], x[1], x[3], x[2])
      if str(x[1]) in freq: 
          task_taskdur = dict()
          task_taskdur = freq[str(x[1])]
          task_taskdur[str(task.id)] = task
          
      else:
          task_taskdur = {str(task.id): task}
          freq[str(x[1])]= task_taskdur
    #Debug: print ('Task #',task.id, len(str(x[0])) ,' name ', task.description, ' duration', task.duration)    
      i = i+1

def weekly_Tasks(week_number,frequency, week_tasks_duration):
    tasks = freq[frequency]
    for task_id in freq[frequency].keys(): 
       week_tasks_duration = week_tasks_duration + tasks[task_id].duration 
    return week_tasks_duration

def frequency_period(frequency):
    match frequency:
        case 'Quinzenal':
            return 1
        case 'Mensal':
            return 3
        case 'Bimestral':
            return 7
        case 'Trimestral':
            return 11
        case 'Semestral':
            return 23
        case 'Anual':
            return 53
    
def alocateTasks(curr_week_number): # allocate tasks according to frequency, duration, place in these order 
    week_tasks_duration = 0
    week_task_list = [] # current week tasks

    for frequency in freq.keys(): #go through all the frequencies
        if frequency == 'Semanal': 
           week_tasks_duration=weekly_Tasks(curr_week_number,frequency, week_tasks_duration)
           week_task_list = list(freq[frequency].keys())         
        else: 
            tasks = []
            tasks = freq[frequency]
            interval = frequency_period(frequency)
            if (curr_week_number - interval) in tasksperWeek: # check prior frequency tasks
              prior_frequency_tasks = tasksperWeek[curr_week_number - interval]
              
              for task_id in freq[frequency].keys(): 
               if  len(prior_frequency_tasks) >0 and task_id in prior_frequency_tasks:  #if task was performed , dont allocate
                  continue
               else:# if task was not performed, allocate concerning to place 
                  curr_task = tasks[task_id]
                  week_tasks_duration = week_tasks_duration + curr_task.duration 
                  if week_tasks_duration  <= MAX:
                     week_task_list.append(task_id)
                     tasksperWeek[curr_week_number]= list(week_task_list)
                  else:
                     week_tasks_duration = week_tasks_duration - curr_task.duration 
    
    tasksperWeek[curr_week_number]= week_task_list
     #   print(' >> ',frequency, ' and dura; ',week_tasks_duration)
    print('<<<<<<<<<<<<<<<<<<<<<<<<< week',curr_week_number, len(tasksperWeek[curr_week_number]), ' tasks' ,'| tasks duration:', round(float(week_tasks_duration)), ' hours') 
    print(tasksperWeek[curr_week_number])
    #week = Week(week_number, week_tasks_duration, week_task_list)

def write_to_excel():
    wb = openpyxl.load_workbook('/Users/maf/Desktop/tarefas2.xlsx')
    mysheet = wb.create_sheet(index=4, title='St')
    i=0
    while i < len( tasksperWeek.keys()):
        print(i)
        i = i+1
        cell='A'+str(i)
        mysheet[cell] = 'Writing new Value!'
        mysheet[cell].value

    wb.save('/Users/maf/Desktop/tarefas2.xlsx')

# readxl returnsa  pylightxl database that holds all worksheets and its data
db = xl.readxl(fn='/Users/maf/Desktop/tarefas.xlsx')

freq = dict() 
tasksperWeek = dict() ##  key: week_number, value: Object Week list of tasks_ids and

storeTasks() 

# get Tasks info - dict{ key: frequency, value: dict{ key: task_id, value: task Object} }
#    current_week = datetime.date(datetime.date.today().year, datetime.date.today().month, datetime.date.today().day).strftime("%V")
i=1
while i <= week_total:
 alocateTasks(i)
 i = i+1
 
write_to_excel()
## conditions:
## frequency is repected
#### one week tasks duration <=8 hours  

# if semanal, bisemanal, and triseanal exists



    