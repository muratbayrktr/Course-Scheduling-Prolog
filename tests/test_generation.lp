% Include the facts from your text files
#include "busy.lp".
#include "courses2.lp".
#include "places.lp".
#include "teaches.lp".
#include "timeslots.lp".

service(Course) :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours), Service = yes.
graduate(Course) :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours), Level = graduate.


{ occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours) 
    : 
        room(Room),
        capacity(Room, MaxCapacity),
        MaxCapacity >= Capacity,
        time_slot(Start),
        time_slot(End),
        End = Start + OccupiedHours,
        teaches(Teacher, Course)
        
} = 1 :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours),
        OccupiedHours = 1..2,
        not service(Course),
        not graduate(Course).




#show occupies/7.
