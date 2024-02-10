#include "busy.lp".
#include "courses.lp".
#include "places.lp".
#include "teaches.lp".
#include "timeslots.lp".

cohort(Y,C) :- course(C,_,_,_,_,_,_,_), Y = (C - 5710000) / 100.

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
        not graduate(Course),
        not ceng100(Course).

{ occupies(5710100, Section, Room, Teacher, Start, End, OccupiedHours) 
    : 
        room(Room),
        capacity(Room, MaxCapacity),
        MaxCapacity >= Capacity,
        time_slot(Start),
        time_slot(End),
        OccupiedHours = Hours,
        End = Start + OccupiedHours,
        teaches(Teacher, 5710100)
        
} = 1 :- course(5710100, Section, Name, Service, Level, Type, Capacity, Hours).

{ occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours) 
    : 
        room(Room),
        capacity(Room, MaxCapacity),
        MaxCapacity >= Capacity,
        time_slot(Start),
        time_slot(End),
        OccupiedHours = Hours,
        End = Start + OccupiedHours,
        teaches(Teacher, Course)
        
} = 1 :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours),
        service(Course),not graduate(Course).  

{ occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours) 
    : 
        room(Room),
        capacity(Room, MaxCapacity),
        MaxCapacity >= Capacity,
        time_slot(Start),
        time_slot(End),
        OccupiedHours = Hours,
        End = Start + OccupiedHours,
        teaches(Teacher, Course)
        
} = 1 :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours),
        not service(Course),graduate(Course).  

{ occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours) 
    : 
        room(Room),
        capacity(Room, MaxCapacity),
        MaxCapacity >= Capacity,
        time_slot(Start),
        time_slot(End),
        OccupiedHours = Hours,
        End = Start + OccupiedHours,
        teaches(Teacher, Course)
        
        
} = 1 :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours),
        service(Course),graduate(Course).  

    
service(Course) :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours), Service = yes.
graduate(Course) :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours), Level = graduate.
ceng100(Course) :- course(Course, Section, Name, Service, Level, Type, Capacity, Hours), Course = 5710100.


same_lecture_same_hour(Course, Section) :-
    occupies(Course, Section, Room1, _, Start1, End1, OccHour1) ,
    occupies(Course, Section, Room2, _, Start2, End2, OccHour2) ,
    Room1 != Room2,
    Start2 >= Start1,
    Start2 < End1.

same_lecture_same_hour(Course, Section) :-
    occupies(Course, Section, Room1, _, Start1, End1, OccHour1) ,
    occupies(Course, Section, Room2, _, Start2, End2, OccHour2) ,
    Room1 != Room2,
    Start1 >= Start2,
    Start1 < End2.

:- occupies(Course, Section, _, _, _, _, _),
    same_lecture_same_hour(Course, Section).


% 2. One classroom cannot be occupied by more than one course at a time.
(Course1, Section1, OccupiedHours1) = (Course2, Section2, OccupiedHours2) :- 
    occupies(Course1, Section1, Room, _, Start1, End1, OccupiedHours1), 
    occupies(Course2, Section2, Room, _, Start2, End2, OccupiedHours2),
    Start2 >= Start1,
    Start2 < End1.

(Course1, Section1, OccupiedHours1) = (Course2, Section2, OccupiedHours2) :- 
    occupies(Course1, Section1, Room, _, Start1, End1, OccupiedHours1), 
    occupies(Course2, Section2, Room, _, Start2, End2, OccupiedHours2),
    Start1 >= Start2,
    Start1 < End2.

% 3. A course cannot be given in a room that is too small.
:- occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours), 
    course(Course,Section, Name, Service, Level, Type, Capacity, OccupiedHours),
    room(Room),
    capacity(Room, MaxCapacity),
    MaxCapacity < Capacity.

%% INSTRUCTORS
% 1. An instructor dont give a course in their busy hour.
conflicting(Course, Section, Room, Teacher, Start, End, OccupiedHours):- 
    occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours), 
    teaches(Teacher, Course),
    busy(Teacher, Hour),
    Hour >= Start,
    Hour < End.

:- occupies(Course, Section, Room, Teacher, Start, End, OccupiedHours),
    conflicting(Course, Section, Room, Teacher, Start, End, OccupiedHours).

% 2. An instructor cannot give more than one course at a time.
:- occupies(Course1, Section1, Room1, Teacher1, Start1, End1, OccupiedHours1), 
    occupies(Course2, Section2, Room2, Teacher2, Start2, End2, OccupiedHours2),
    Course1 != Course2,
    Teacher1 = Teacher2,
    Start2 >= Start1,
    Start2 < End1.

:- occupies(Course1, Section1, Room1, Teacher1, Start1, End1, OccupiedHours1), 
    occupies(Course1, Section2, Room2, Teacher2, Start2, End2, OccupiedHours2),
    Section1 != Section2,
    Teacher1 = Teacher2,
    Start2 >= Start1,
    Start2 < End1.

:- occupies(Course1, Section1, Room1, Teacher1, Start1, End1, OccupiedHours1), 
    occupies(Course1, Section1, Room2, Teacher2, Start2, End2, OccupiedHours2),
    OccupiedHours1 != OccupiedHours2,
    Teacher1 = Teacher2,
    Start2 >= Start1,
    Start2 < End1.

:- occupies(Course, 1, Room1, Teacher, Start1, End1, OccupiedHours1), 
    occupies(Course, 2, Room2, Teacher, Start2, End2, OccupiedHours2),
    Room1 != Room2,
    Start2 >= Start1,
    Start2 < End1.


%% 

%% COURSES
% 3. If a non-service undergraduate level course has splitted into three sections, one hour of lecture must be conducted at BMB1 at least
:- occupies(Course, _, Room1, _, _, _, 1),
    occupies(Course, _, Room2, _, _, _, 2),
    not service(Course),
    not graduate(Course),
    Room1 != bmb1,
    Room2 != bmb1,
    occupies(Course, 1, _, _, _, _, _),
    occupies(Course, 2, _, _, _, _, _),
    occupies(Course, 3, _, _, _, _, _).


% 4. If a non-service undergraduate level course has splitted into two sections, all hours must be conducted at BMB1 at least
:- occupies(Course, _, Room1, _, _, _, 1),
    occupies(Course, _, Room2, _, _, _, 2),
    not service(Course),
    not graduate(Course),
    occupies(Course, 1, _, _, _, _, _),
    occupies(Course, 2, _, _, _, _, _),
    not occupies(Course, 3, _, _, _, _, _),
    all_different_than_bmb1(Room1, Room2).

all_different_than_bmb1(Room1, Room2) :-
    room(Room1), room(Room2),
    Room1 != bmb1.

all_different_than_bmb1(Room1, Room2) :-
    room(Room2), room(Room1),
    Room2 != bmb1.

% 5. If a course taught by at least two instructor, if one of them teach to a section, other instructor must not teach to that section.

% Constraint: In courses with multiple instructors, each section must be taught by only one of them.
num_instructors(Course, Count) :-
    course(Course, _, _, _, _, _, _, _),
    Count = #count { Teacher : teaches(Teacher, Course) }.

num_sections(Course, Count) :-
    course(Course, _, _, _, _, _, _, _),
    Count = #count { Section : occupies(Course, Section, _, _, _, _, _) }.


%% 2 SECTIONS CASE
% If sections are same, teachers can't be different, this is grounded.
:- occupies(Course, Section1, _, Teacher1, _, _, _),
    occupies(Course, Section1, _, Teacher2, _, _, _),
    num_instructors(Course, Count1),
    num_sections(Course, Count2),
    Count1 = Count2,
    Teacher1 != Teacher2.

% If sections are different, teachers must be different
:- occupies(Course, Section1, _, Teacher1, _, _, _),
    occupies(Course, Section2, _, Teacher2, _, _, _),
    num_instructors(Course, Count1),
    num_sections(Course, Count2),
    Count1 = Count2,
    Section1 != Section2,
    Teacher1 = Teacher2.

%% 3 SECTIONS CASE
% If sections are same, teachers can't be different, this is grounded. (INTRA-SECTION CONSISTENCY)
:- occupies(Course, Section1, _, Teacher1, _, _, _),
    occupies(Course, Section1, _, Teacher2, _, _, _),
    num_instructors(Course, Count1),
    num_sections(Course, Count2),
    Count1 = 2,
    Count2 = 3,
    Teacher1 != Teacher2.



% If sections are different, teachers must be different (INTER-SECTION CONSISTENCY)
:- occupies(Course, 1, _, Teacher1, _, _, _),
    occupies(Course, 2, _, Teacher2, _, _, _),
    occupies(Course, 3, _, Teacher3, _, _, _),
    num_instructors(Course, Count1),
    num_sections(Course, Count2),
    Count1 = 2,
    Count2 = 3,
    
    Teacher1 = Teacher2,
    Teacher1 = Teacher3,
    Teacher2 = Teacher3.

%% Courses in the same cohort cannot be given at the same time (~).
:- occupies(Course1, Section1, _, _, Start1, End1, _),
    occupies(Course2, Section2, _, _, Start2, End2, _),
    course(Course1, Section1, _, no, undergraduate, must, _, _),
    course(Course2, Section2, _, no, undergraduate, must, _, _),
    cohort(Cohort1, Course1),
    cohort(Cohort2, Course2),
    Course1 != Course2,
    Cohort1 = Cohort2,
    Start2 >= Start1,
    Start2 < End1.



%% Same course, same section different hour time difference should be greater than 38
:- occupies(Course, Section, _, _, Start1, End1, OccHour1),
    occupies(Course, Section, _, _, Start2, End2, OccHour2),
    OccHour1 != OccHour2,
    Start2 >= End1,
    Start2 < End1 + 41.



%% Gaps between same cohorts should be minimized (~). The penalty applied is 5 points for each hour of gap.
gap_count(Course1, Section1, OccHour1, Course2, Section2, OccHour2, GapCount) :-
    occupies(Course1, Section1, _, _, Start1, End1, OccHour1),
    occupies(Course2, Section2, _, _, Start2, End2, OccHour2),
    course(Course1, Section1, _, no, undergraduate, must, _, _),
    course(Course2, Section2, _, no, undergraduate, must, _, _),
    cohort(Cohort1, Course1),
    cohort(Cohort2, Course2),
    Cohort1 = Cohort2,
    Start2 > End1,
    GapCount = Start2 - End1.



:~ gap_count(Course1, Section1, OccHour1, Course2, Section2, OccHour2, GapCount). [GapCount*5]

#show occupies/7.
