   #pragma once  
   #include <xlnt/xlnt.hpp>  
   #include <string>  
   #include <mutex>  
   #include <fstream>  

   ///singleton Class  
   class Data_Base  
   {  
      //Fields  
   private:  
      xlnt::workbook Gym_Padel_DB;  
      xlnt::worksheet User;  
      xlnt::worksheet Manager;  
      xlnt::worksheet Trainee;  
      xlnt::worksheet HallSystem;  
      xlnt::worksheet Class;  
      xlnt::worksheet StaffMember;  
      xlnt::worksheet Coach;  
      xlnt::worksheet Receptionist;  
      xlnt::worksheet Subscription;  
      xlnt::worksheet WorkoutPlan;  
      xlnt::worksheet TraineeAssignedToClass;  
      xlnt::worksheet CoachTrainingClasses;  
      xlnt::worksheet HallHoldsClasses;  
      xlnt::worksheet workoutPlanAssignedToClass;  
      xlnt::worksheet Manages;  
      xlnt::worksheet WorksIn;  
      static Data_Base* InstanceDB;  
      static std::mutex mtx;  
      Data_Base();  

   public:  
      Data_Base(Data_Base& t) = delete;  
      static Data_Base* GetDB();  
      std::ofstream fileee;  

      xlnt::workbook getGym_Padel_DB(std::string FileName);  
      void SaveGym_Padel_DB(std::string FileName);  
      xlnt::worksheet GetUser(std::string FileName);  
      xlnt::worksheet GetManager(std::string FileName);  
      xlnt::worksheet GetTrainee(std::string FileName);  
      xlnt::worksheet GetCoach(std::string FileName);  
      xlnt::worksheet GetClass(std::string FileName);  
      xlnt::worksheet GetHallSystem(std::string FileName);  
      xlnt::worksheet GetWorkoutPlan(std::string FileName);  
      xlnt::worksheet GetReceptionist(std::string FileName);  
      xlnt::worksheet GetSubscription(std::string FileName);  
      xlnt::worksheet GetAssignedTo(std::string FileName);  
      xlnt::worksheet GetManages(std::string FileName);  
      xlnt::worksheet GetWorksIn(std::string FileName);  
      xlnt::worksheet GetHallHoldsClasses(std::string FileName);  
      xlnt::worksheet GetworkoutPlanAssignedToClass(std::string FileName);  
      xlnt::worksheet GetCoachTrainingClasses(std::string FileName);  
      xlnt::worksheet GetStaffMember(std::string FileName);  
   };
