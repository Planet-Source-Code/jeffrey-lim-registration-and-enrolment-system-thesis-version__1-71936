
-----------------------------------------------------------
-- Dclc database - CREATE
-- Thesis Group Name

-- this script will drop an existing database 
-- and create a fresh new installation
-----------------------------------------------------------
-- Drop and Create Database

USE master
GO
IF EXISTS (SELECT * FROM SysDatabases WHERE NAME='Dclc') DROP DATABASE Dclc
go

-- This creates the database data file and log file on the default directories
CREATE DATABASE Dclc
go

USE Dclc
go

-----------------------------------------------------------
-- Create Tables, in order from primary to secondary
CREATE TABLE dbo.Courses (
  CourseCode			VARCHAR(6) NOT NULL PRIMARY KEY NONCLUSTERED,
  CourseDesc			VARCHAR(100),
  College				VARCHAR(20),
  CourseYear			INT,
  Deleted			BIT DEFAULT 0
  );
-- College: Nursing, Engineering
-- CourseYear: 1, 2, 3, 4 or 5 year course
go 

CREATE TABLE dbo.Sy (
  SyId				INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  Sy				VARCHAR(9) NOT NULL,
  );
-- Sy: 2008-2009
go

CREATE TABLE dbo.YearLevel (
  YearLevelId			INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  YearLevel				VARCHAR(12) NOT NULL,
  Deleted			BIT DEFAULT 0
  );
-- YearLevel: 1st Year, 2nd Year, 3rd Year, 4th Year, 5th Year
--            nb: First Char must be numeric
go

CREATE TABLE dbo.Semester (
  SemesterId			INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  Semester				VARCHAR(15) NOT NULL,
  Deleted			BIT DEFAULT 0
  );
-- Semester: 1st Semester, 2nd Semester, Summer
go

IF EXISTS (SELECT TABLE_NAME FROM Information_Schema.Tables WHERE table_name='Students') DROP TABLE Students
go
CREATE TABLE dbo.Students (
  StudentId					VARCHAR(10) NOT NULL PRIMARY KEY NONCLUSTERED,
  Firstname					VARCHAR(35) NOT NULL,
  Lastname					VARCHAR(35) NOT NULL,
  Middlename					VARCHAR(35),
  Gender					VARCHAR(1),
  BirthDate					DATETIME NULL,
  Address					VARCHAR(100),
  Nationality					VARCHAR(20),
  Religion					VARCHAR(20),
  CourseCode					VARCHAR(6) NOT NULL FOREIGN KEY REFERENCES dbo.Courses,
  YearLevelId					INT NOT NULL FOREIGN KEY REFERENCES dbo.YearLevel,
  SemesterId					INT NOT NULL FOREIGN KEY REFERENCES dbo.Semester,
  LastSchoolAttended				VARCHAR(60),
  DateEnrolled					DATETIME NULL,
  Status					VARCHAR(12) DEFAULT 'Active'
  );
-- Status: Active, Graduate, Blacklisted, Dean's lister, ...
go 

CREATE TABLE dbo.Credentials (
  StudentId			VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Students,
  Form137			BIT,
  Form138			BIT, 
  Gmrc				BIT,
  BirthCertificate	BIT,
  HsDiploma			BIT 
  );
-- (orig) = Form137, Form138, Gmrc
-- (xcopy) = BirthCert, Hs Diploma
go 

CREATE TABLE dbo.Parents (
  StudentId			VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Students,
  Father			VARCHAR(45),
  FatherOccupation	VARCHAR(20),
  Mother			VARCHAR(45),
  MotherOccupation	VARCHAR(20),
  Address			VARCHAR(100)
  );
go 

-- CREATE CLUSTERED INDEX IxUsersName ON dbo.Users (Username, Username);
-- ALTER TABLE dbo.Users ADD CONSTRAINT FK_Users_Father FOREIGN KEY	(UserId) REFERENCES dbo.Users (UsersId);
-- ALTER TABLE dbo.Users ADD CONSTRAINT FK_Users_Mother FOREIGN KEY	(UsersId) REFERENCES dbo.Users (UsersId);
-- go 

CREATE TABLE dbo.Users (
  UserId		VARCHAR(10) NOT NULL PRIMARY KEY,
  Username		VARCHAR(35) NOT NULL,
  Password		VARCHAR(20) NOT NULL,
  UserType		VARCHAR(15) NULL,
  UserStatus	VARCHAR(1) DEFAULT 'A'
  );
go 

CREATE TABLE dbo.UserLog (
  UserId		VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Users,
  [Login]		DATETIME NULL,
  LogOut		DATETIME NULL
  );
go 

CREATE TABLE dbo.SchoolInfo (
  SchoolName1			VARCHAR(30),
  SchoolName2			VARCHAR(30),
  Address				VARCHAR(60),
  TelNo1				VARCHAR(15),
  TelNo2				VARCHAR(15),
  FaxNo					VARCHAR(15),
  CurrentSyId			INT NOT NULL FOREIGN KEY REFERENCES dbo.Sy,
  CurrentSemesterId		INT NOT NULL FOREIGN KEY REFERENCES dbo.Semester
  );
go 

CREATE TABLE dbo.LastNo (
  StudentNo			INT,
  ReceiptNo			INT,
  AssessNo			INT
  );
go

-- Accounting
CREATE TABLE dbo.Fees (
  FeesId				INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  CourseCode				VARCHAR(6) NOT NULL FOREIGN KEY REFERENCES dbo.Courses,
  CurrentSyId				INT NOT NULL FOREIGN KEY REFERENCES dbo.Sy,
  CurrentSemesterId			INT NOT NULL FOREIGN KEY REFERENCES dbo.Semester,
  CurrentYearLevelId				INT NOT NULL FOREIGN KEY REFERENCES dbo.YearLevel,
-- ..
  Entrance					NUMERIC(12,2) DEFAULT 0,
  TuitionFee				NUMERIC(12,2) DEFAULT 0,
-- Others
  Registration				NUMERIC(12,2) DEFAULT 0,
  Library					NUMERIC(12,2) DEFAULT 0,
  Laboratory				NUMERIC(12,2) DEFAULT 0,
  AthleticFee				NUMERIC(12,2) DEFAULT 0,
  GuidanceAndCounselor		NUMERIC(12,2) DEFAULT 0,
-- Misc
  Rle						NUMERIC(12,2) DEFAULT 0,
  Affiliation						NUMERIC(12,2) DEFAULT 0,
  NursingAudit				NUMERIC(12,2) DEFAULT 0,
  MarineLaboratory						NUMERIC(12,2) DEFAULT 0,
  SpeechLab					NUMERIC(12,2) DEFAULT 0,
  HrmLab					NUMERIC(12,2) DEFAULT 0,
  Ojt						NUMERIC(12,2) DEFAULT 0,
  Rta						NUMERIC(12,2) DEFAULT 0,
  HOn						NUMERIC(12,2) DEFAULT 0,
  Mta						NUMERIC(12,2) DEFAULT 0,
  IdNamePlate				NUMERIC(12,2) DEFAULT 0,
  Sdf						NUMERIC(12,2) DEFAULT 0,
  PowerFee					NUMERIC(12,2) DEFAULT 0,
  Internet					NUMERIC(12,2) DEFAULT 0,
  Internship				NUMERIC(12,2) DEFAULT 0,
  Waiver					NUMERIC(12,2) DEFAULT 0,
  Nstp						NUMERIC(12,2) DEFAULT 0,
-- Total
  TotalCashBasis			NUMERIC(12,2) DEFAULT 0,
  TotalInstallmentBasis		NUMERIC(12,2) DEFAULT 0,
  DownPayment				NUMERIC(12,2) DEFAULT 0,
-- Deleted Marker
  Deleted					BIT DEFAULT 0
  );
go 

CREATE TABLE dbo.Ledger (
  LedgerId				INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  ReceiptNo				VARCHAR(8),
  StudentId				VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Students,
  SyId					INT NOT NULL FOREIGN KEY REFERENCES dbo.Sy,
  YearLevelId			INT NOT NULL FOREIGN KEY REFERENCES dbo.YearLevel,
  SemesterId			INT NOT NULL FOREIGN KEY REFERENCES dbo.Semester,
  Debit					NUMERIC(12,2) DEFAULT 0, --pautang
  Credit				NUMERIC(12,2) DEFAULT 0, --utang
  Particular			VARCHAR(25) NULL,
  TranDate				DATETIME NULL,
  PostedBy				VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Users(UserId)
  );
go 


-- Semesters
INSERT dbo.Semester (Semester) VALUES ('1st Semester');
INSERT dbo.Semester (Semester) VALUES ('2nd Semester');
INSERT dbo.Semester (Semester) VALUES ('Summer');
go

-- School Year
INSERT dbo.Sy (Sy) VALUES ('2008-2009');
go

-- Year Level
INSERT dbo.YearLevel (YearLevel) VALUES ('1st Year');
INSERT dbo.YearLevel (YearLevel) VALUES ('2nd Year');
INSERT dbo.YearLevel (YearLevel) VALUES ('3rd Year');
INSERT dbo.YearLevel (YearLevel) VALUES ('4th Year');
INSERT dbo.YearLevel (YearLevel) VALUES ('5th Year');
go

-- Year Level
INSERT dbo.Courses (CourseCode, CourseDesc, College, CourseYear) VALUES ('BSN', 'Bachelor of Science in Nursing', 'Nursing', 4);
INSERT dbo.Courses (CourseCode, CourseDesc, College, CourseYear) VALUES ('BSBA', 'Bachelor of Science in Business Administration Major in Financial Management', 'Management', 4);
INSERT dbo.Courses (CourseCode, CourseDesc, College, CourseYear) VALUES ('BSMT', 'Bachelor of Science in Marine Transportation', 'Marine', 4);
INSERT dbo.Courses (CourseCode, CourseDesc, College, CourseYear) VALUES ('BSCS', 'Bachelor of Science in Computer Science', 'ComSci', 4);


-- Execute Dependency Here
CREATE TABLE dbo.Rooms (RoomId			INT IDENTITY NOT NULL PRIMARY KEY NONCLUSTERED,
                        RoomNo			VARCHAR(6) NOT NULL, -- N-101, E-202
						FloorLevel			VARCHAR(4), -- 1st, Second...
						RoomType			VARCHAR(4), -- Lab/Lec
						Building			VARCHAR(12)); -- Science 
go
CREATE TABLE dbo.Instructors (InstructorId   INT IDENTITY NOT NULL PRIMARY KEY NONCLUSTERED,
                              Instructor		VARCHAR(45));
go
CREATE TABLE dbo.Subjects (SubjectCode   VARCHAR(10) NOT NULL PRIMARY KEY NONCLUSTERED,
                              SubjectDesc		VARCHAR(50),
							  LecUnits			NUMERIC(2,1) DEFAULT 0,
							  LabUnits			NUMERIC(2,1) DEFAULT 0); 
go

IF EXISTS (SELECT TABLE_NAME FROM Information_Schema.Tables WHERE table_name='Schedules') DROP TABLE Schedules
go
CREATE TABLE dbo.Schedules (
  SchedCode			INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  SubjectCode		VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Subjects,
  CourseCode		VARCHAR(6) FOREIGN KEY REFERENCES dbo.Courses,
  TimeSchedStart	VARCHAR(8),
  TimeSchedEnd		VARCHAR(8),
  DaysSched			VARCHAR(10), -- MWF
  SyId				INT NOT NULL FOREIGN KEY REFERENCES dbo.Sy,
  SemesterId		INT NOT NULL FOREIGN KEY REFERENCES dbo.Semester,
  RoomId			INT NOT NULL FOREIGN KEY REFERENCES dbo.Rooms,
  InstructorId		INT NOT NULL FOREIGN KEY REFERENCES dbo.Instructors
  );
GO 
-- Sample Dependency Data
-- ROOMS
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('A-608', '6th', 'Lec', 'Science');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('A-706', '7th', 'Lec', 'Science');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('S-101', '1st', 'Lec', 'Science');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('VT-202', '2nd', 'Lab', 'East Wing');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('M-303', '3rd', 'Lec', 'Main');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('M-301', '3rd', 'Lec', 'Main');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('VT-201', '2nd', 'Lab', 'East Wing');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('GYM', '1st', 'Lab', 'BA12A');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('POOL', '1st', 'Lab', 'BA12A');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('CL', '1st', 'Lab', 'BA12A');
INSERT dbo.Rooms (RoomNo, FloorLevel, RoomType, Building) VALUES('TR', '1st', 'Lab', 'BA12A');
-- INSTRUCTORS
INSERT dbo.Instructors (Instructor) VALUES('Jack Bauer');
INSERT dbo.Instructors (Instructor) VALUES('Tony Almeida');
INSERT dbo.Instructors (Instructor) VALUES('Chloe O`Brian');
INSERT dbo.Instructors (Instructor) VALUES('Michael Scofield');
INSERT dbo.Instructors (Instructor) VALUES('Lincoln Burrows');
-- SUBJECTS
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Chem1', 'General Chemistry (Organic and Inorganic)', 3, 2);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Physics', 'Physics', 2, 1);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('NuDiet', 'Nutrition and Diet Therapy', 3, 1);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Ncm105', 'Related Learning Experiences', 0, 6);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Ncm107', 'Nursing Leadership And Management', 8, 8);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('AI', 'Artificial Intelligence', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('DBMS', 'Database Management System', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('SAD', 'System Analysis And Design', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Elec', 'Computer Elective', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Automt', 'Automata & Language Theory', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('Sofend', 'Software Engineering', 3, 0);
INSERT dbo.Subjects (SubjectCode, SubjectDesc, LecUnits, LabUnits) VALUES('PE1', 'Rythmic, Folk', 2, 0);
-- SCHEDULES
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('Chem1', 'BSN', '8:00', '9:00', 'MWF', 1, 1, 5, 1);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('Physics', 'BSN', '9:00', '10:00', 'TTh', 1, 1, 3, 2);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('NuDiet', 'BSN', '10:00', '11:00', 'S', 1, 1, 1, 3);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('AI', 'BSCS', '11:00', '12:00', 'MW', 1, 1, 2, 4);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('DBMS', 'BSCS', '12:00', '13:00', 'F', 1, 1, 4, 5);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('SAD', 'BSCS', '13:00', '14:00', 'TTh', 1, 1, 5, 1);

INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('PE1', 'BSCS', '15:00', '16:00', 'M', 1, 1, 5, 1);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('PE1', 'BSCS', '16:00', '17:00', 'W', 1, 1, 5, 1);
INSERT dbo.Schedules (SubjectCode, CourseCode, TimeSchedStart, TimeSchedEnd, DaysSched, SyId, SemesterId, RoomId, InstructorId) 
 VALUES('PE1', 'BSCS', '17:00', '18:00', 'F', 1, 1, 5, 1);
-- eo: Dependency Execution



-- Finally, Enrollment System
IF EXISTS (SELECT TABLE_NAME FROM Information_Schema.Tables WHERE table_name='Enrolled') DROP TABLE Enrolled
go
CREATE TABLE dbo.Enrolled (
  EnrollId			INT NOT NULL IDENTITY PRIMARY KEY NONCLUSTERED,
  AssessNo			INT NOT NULL,
  StudentId			VARCHAR(10) NOT NULL FOREIGN KEY REFERENCES dbo.Students,
  SchedCode			INT NOT NULL FOREIGN KEY REFERENCES dbo.Schedules,
  FeesId			INT NOT NULL FOREIGN KEY REFERENCES dbo.Fees,
  CashBasis			BIT DEFAULT 0,
  Status			VARCHAR(1) DEFAULT 'P' -- P=Pending, C-Confirmed
  );
go 


-- DCLC-RES Default Data
-- Default Users
INSERT dbo.users (UserId, Username, Password, UserType) VALUES('admin', 'Administrator', 'admin', 'Administrator');
INSERT dbo.users (UserId, Username, Password, UserType) VALUES('cashier', 'Cashier', 'cashier', 'Cashier');
INSERT dbo.users (UserId, Username, Password, UserType) VALUES('registrar', 'Registrar', 'registrar', 'Registrar');
go

-- School Info
INSERT dbo.SchoolInfo (SchoolName1, SchoolName2, Address, TelNo1, TelNo2, FaxNo, CurrentSyId, CurrentSemesterId) VALUES ('DR. CARLOS S. LANTING COLLEGE', 'CASUAL GENERAL HOSPITAL', '16 Tandang Sora Ave. Sangandaan, Quezon City, Philippines', '(02) 938-7782', '(02)938-7789', '(02)939-7229', '1', '1');

-- Initialize LastNo
TRUNCATE TABLE dbo.LastNo;
INSERT INTO dbo.LastNo(StudentNo, ReceiptNo, AssessNo) VALUES(0, 0, 0);

SELECT * FROM users;
SELECT * FROM Semester;
SELECT * FROM Sy;
SELECT * FROM LastNo;





