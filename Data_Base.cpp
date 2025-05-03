#include "Data_Base.h"

Data_Base* Data_Base::InstanceDB = nullptr;
std::mutex Data_Base::mtx;

Data_Base* Data_Base:: GetDB() {
	if (InstanceDB == nullptr) {
		std::lock_guard<std::mutex>lock(mtx);//lock multithreaded
		if (InstanceDB == nullptr)
			InstanceDB = new Data_Base();
	}
	return InstanceDB;
}

Data_Base::Data_Base() {
		
}


xlnt::workbook Data_Base:: getGym_Padel_DB(std::string FileName) {
	try
	{
		this->Gym_Padel_DB.load(FileName);
	}
	catch (const std::exception&)
	{
		std::cout << "File Not Exist\n";
	}

	return this->Gym_Padel_DB;
}
//************************************************\\

xlnt::worksheet Data_Base:: GetUser(std::string FileName) {
	try
	{
		this->User = this->Gym_Padel_DB.sheet_by_index(0);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->User = this->Gym_Padel_DB.sheet_by_index(0);
		this->User.cell("A1").value("ID");
		this->User.cell("B1").value("Name");
		this->User.cell("C1").value("Email");
		this->User.cell("D1").value("Phone Number");
		this->User.cell("E1").value("Password");
		this->User.cell("F1").value("Date Of Birth");
		this->User.cell("G1").value("Gender");
		this->User.cell("H1").value("Gym Or Padel");
		this->User.cell("H1").value("Assigned");
	}
	return this->User;
}

//************************************************\\
}


xlnt::worksheet Data_Base:: GetManager(std::string FileName) {
	try
	{
		this->Manager = this->Gym_Padel_DB.sheet_by_index(1);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Manager = this->Gym_Padel_DB.sheet_by_index(1);
		this->Manager.cell("A1").value("ID");

	}
	return this->Manager;
}




//************************************************\\
	

xlnt::worksheet Data_Base:: GetTrainee(std::string FileName) {
	try
	{
		this->Trainee = this->Gym_Padel_DB.sheet_by_index(2);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Trainee = this->Gym_Padel_DB.sheet_by_index(2);
		this->Trainee.cell("A1").value("ID");
		this->Trainee.cell("B1").value("Weight");
		this->Trainee.cell("C1").value("Height");
		this->Trainee.cell("D1").value("Lost Calories Per Minute");
	}

	return this->Trainee;
}

//************************************************\\
	



xlnt::worksheet Data_Base:: GetCoach(std::string FileName) {
	try
	{
		this->Coach = this->Gym_Padel_DB.sheet_by_index(3);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Coach = this->Gym_Padel_DB.sheet_by_index(3);
		this->Coach.cell("A1").value("ID");
		this->Coach.cell("B1").value("Weight");
		this->Coach.cell("C1").value("Height");
		this->Coach.cell("D1").value("Years Of Experience");
	}
	return this->Coach;
}


//************************************************\\



xlnt::worksheet Data_Base:: GetClass(std::string FileName) {
	try
	{
		this->Class = this->Gym_Padel_DB.sheet_by_index(4);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Class = this->Gym_Padel_DB.sheet_by_index(4);
		this->Class.cell("A1").value("Class Code");
		this->Class.cell("B1").value("Gym Or Padel");
		this->Class.cell("C1").value("Training Days");
		this->Class.cell("D1").value("Start Time");
		this->Class.cell("E1").value("End Time");
		this->Class.cell("F1").value("Max Number Of Trainees");
	}
	return this->Class;
}
//************************************************\\
	


xlnt::worksheet Data_Base:: GetHallSystem(std::string FileName) {
	try
	{
		this->HallSystem = this->Gym_Padel_DB.sheet_by_index(5);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->HallSystem = this->Gym_Padel_DB.sheet_by_index(5);
		this->HallSystem.cell("A1").value("Hall Code");
		this->HallSystem.cell("B1").value("Gym Or Padel");
		this->HallSystem.cell("C1").value("Hall Capacity");
		this->HallSystem.cell("D1").value("Working Days");
		this->HallSystem.cell("E1").value("Start Time");
		this->HallSystem.cell("F1").value("End Time");
	}
	return this->HallSystem;
}

//************************************************\\



xlnt::worksheet Data_Base:: GetWorkoutPlan(std::string FileName) {
	try
	{
		this->WorkoutPlan = this->Gym_Padel_DB.sheet_by_index(6);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->WorkoutPlan = this->Gym_Padel_DB.sheet_by_index(6);
		this->WorkoutPlan.cell("A1").value("Workout Plan Code");
		this->WorkoutPlan.cell("B1").value("Intensity");
		this->WorkoutPlan.cell("C1").value("Duration");
	}
	return this->HallSystem;
}


//************************************************\\
	

xlnt::worksheet Data_Base:: GetReceptionist(std::string FileName) {
	try
	{
		this->Receptionist = this->Gym_Padel_DB.sheet_by_index(7);
	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Receptionist = this->Gym_Padel_DB.sheet_by_index(7);

		this->Receptionist.cell("A1").value("ID");
	}
	return this->Receptionist;
}
//************************************************\\
	

xlnt::worksheet Data_Base:: GetSubscription(std::string FileName) {
	try
	{
		this->Subscription = this->Gym_Padel_DB.sheet_by_index(8);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Subscription = this->Gym_Padel_DB.sheet_by_index(8);
		this->Subscription.cell("A1").value("ID");
		this->Subscription.cell("B1").value("Start Date");
		this->Subscription.cell("C1").value("End Date");
		this->Subscription.cell("D1").value("Price");
		this->Subscription.cell("E1").value("Type");
	}
	return this->Subscription;
}


//************************************************\\
	

xlnt::worksheet Data_Base:: GetAssignedTo(std::string FileName) {
	try
	{
		this->TraineeAssignedToClass = this->Gym_Padel_DB.sheet_by_index(9);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->TraineeAssignedToClass = this->Gym_Padel_DB.sheet_by_index(9);
		this->TraineeAssignedToClass.cell("A1").value("Manager ID");
		this->TraineeAssignedToClass.cell("B1").value("Hall Code");
	}
	return this->TraineeAssignedToClass;
}
//************************************************\\

xlnt::worksheet Data_Base:: GetManages(std::string FileName) {
	try
	{
		this->Manages = this->Gym_Padel_DB.sheet_by_index(10);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->Manages = this->Gym_Padel_DB.sheet_by_index(10);
		this->Manages.cell("A1").value("Manager ID");
		this->Manages.cell("B1").value("Hall Code");
	}
	return this->Manages;
}
//************************************************\\


xlnt::worksheet Data_Base:: GetWorksIn(std::string FileName) {
	try
	{
		this->WorksIn = this->Gym_Padel_DB.sheet_by_index(11);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->WorksIn = this->Gym_Padel_DB.sheet_by_index(11);
		this->WorksIn.cell("A1").value("Receptionist ID");
		this->WorksIn.cell("B1").value("Hall Code");
	}
	return this->WorksIn;
}

//************************************************\\
	


xlnt::worksheet Data_Base:: GetHallHoldsClasses(std::string FileName) {
	try
	{
		this->HallHoldsClasses = this->Gym_Padel_DB.sheet_by_index(12);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->HallHoldsClasses = this->Gym_Padel_DB.sheet_by_index(12);
		this->HallHoldsClasses.cell("A1").value("Hall Code");
		this->HallHoldsClasses.cell("B1").value("Class Code");
	}
	return this->HallHoldsClasses;
}

//************************************************\\
	


xlnt::worksheet Data_Base:: GetworkoutPlanAssignedToClass(std::string FileName) {
	try
	{
		this->workoutPlanAssignedToClass = this->Gym_Padel_DB.sheet_by_index(13);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->workoutPlanAssignedToClass = this->Gym_Padel_DB.sheet_by_index(13);
		this->workoutPlanAssignedToClass.cell("A1").value("Workout Plan Code");
		this->workoutPlanAssignedToClass.cell("B1").value("Class Code");
	}
	return this->workoutPlanAssignedToClass;
}
//************************************************\\
	


xlnt::worksheet Data_Base:: GetCoachTrainingClasses(std::string FileName) {
	try
	{
		this->CoachTrainingClasses = this->Gym_Padel_DB.sheet_by_index(14);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->CoachTrainingClasses = this->Gym_Padel_DB.sheet_by_index(14);
		this->CoachTrainingClasses.cell("A1").value("Coach ID");
		this->CoachTrainingClasses.cell("B1").value("Class Code");
	}
	return this->CoachTrainingClasses;
}
//************************************************\\
	


xlnt::worksheet Data_Base:: GetStaffMember(std::string FileName) {
	try
	{
		this->StaffMember = this->Gym_Padel_DB.sheet_by_index(15);

	}
	catch (const std::exception&)
	{
		this->Gym_Padel_DB.create_sheet().title(FileName);
		this->StaffMember = this->Gym_Padel_DB.sheet_by_index(15);
		this->StaffMember.cell("A1").value("ID");
		this->StaffMember.cell("B1").value("Net Salary");
		this->StaffMember.cell("C1").value("Start Time");
		this->StaffMember.cell("D1").value("End Time");
		this->StaffMember.cell("E1").value("Type");
		this->StaffMember.cell("F1").value("Working Days");


	}
	return this->StaffMember;
}

