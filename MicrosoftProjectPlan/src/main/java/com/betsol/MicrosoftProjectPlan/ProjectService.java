package com.betsol.MicrosoftProjectPlan;

import java.math.BigDecimal;
import java.util.Date;
import java.util.Iterator;

import com.aspose.tasks.*;
import com.betsol.MicrosoftProjectPlan.AddLicense;
import com.betsol.MicrosoftProjectPlan.Utils;

public class ProjectService {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ProjectService.class);

		AddLicense.add();

		long OneSec = 10000000;// microsecond * 10
		long OneMin = 60 * OneSec;
		long OneHour = 60 * OneMin;

		// Read an existing project
		Project project = new Project(dataDir + "Project.mpp");
		/*
		 * Iterator <Task> taskItr =
		 * project.getRootTask().getChildren().iterator();
		 * 
		 * while(taskItr.hasNext()){ Task task = taskItr.next(); RemoveTask
		 * rTask = new RemoveTask(task); rTask.preAlg(task, 0); task.delete();
		 * break; }
		 */
		Date date = new Date();
		Iterator<Task> taskItr = project.getRootTask().getChildren().iterator();
		int count = 0;
		while (taskItr.hasNext()) {
			Task task = taskItr.next();
			// cal.set(2016, 8, 1, 8, 0,0);
			task.set(Tsk.START, new Date(date.getTime() + 8*24*60*60*1000));

			// cal.set(2016, 8, 7, 17, 0, 0);
			task.set(Tsk.FINISH, new Date(date.getTime() + 18*24*60*60*1000));
			project.getResources().add("RES1");
			project.getResourceAssignments().add(task, project.getResources().getById(0));

			++count;
		}

		// create a new task
		date = new Date();
		Task task = project.getRootTask().getChildren().add("Task" + ++count);
		task.set(Tsk.START, new Date(date.getTime() + 19*24*60*60*1000));
		task.set(Tsk.FINISH, new Date(date.getTime() + 29*24*60*60*1000));
		Task subtask = task.getChildren().add("Subtask1");
		project.getResources().add("RES2");
		project.getResourceAssignments().add(task, project.getResources().getById(1));

		// create a new task
		date = new Date();
		Task task1 = project.getRootTask().getChildren().add("Task" + ++count);
		task1.set(Tsk.START, new Date(date.getTime() + 30*24*60*60*1000));
		task1.set(Tsk.FINISH, new Date(date.getTime() + 39*24*60*60*1000));
		subtask = task1.getChildren().add("Subtask1");
		subtask = task1.getChildren().add("Subtask2");
		subtask = task1.getChildren().add("Subtask3");
		subtask = task1.getChildren().add("Subtask4");
		task1.set(Tsk.DURATION, project.getDuration(2));
		task1.set(Tsk.PERCENT_COMPLETE, 50);
		task1.set(Tsk.COST, BigDecimal.valueOf(800));
		project.getResources().add("RES3");
		project.getResourceAssignments().add(task1, project.getResources().getById(2));

		// create a new task
		date = new Date();
		Task task2 = project.getRootTask().getChildren().add("Task" + ++count);
		task2.set(Tsk.START, new Date(date.getTime() + 31*24*60*60*1000));
		task2.set(Tsk.FINISH, new Date(date.getTime() + 40*24*60*60*1000));
		project.getResources().add("RES4");
		project.getResourceAssignments().add(task2, project.getResources().getById(3));

		// create a new task
		date = new Date();
		Task task3 = project.getRootTask().getChildren().add("Task" + ++count);
		task3.set(Tsk.START, new Date(date.getTime() + 41*24*60*60*1000));
		task3.set(Tsk.FINISH, new Date(date.getTime() + 49*24*60*60*1000));
		project.getResources().add("RES5");
		project.getResourceAssignments().add(task3, project.getResources().getById(4));

		Date date1 = new Date();
		// Save the project as MPP project file
		project.save(dataDir + "ProjectPlanOutput-" + date1.getTime() + ".mpp", SaveFileFormat.MPP);
		// Display result of conversion.
		System.out.println("Process completed Successfully");
	}
}
