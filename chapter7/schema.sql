CREATE TABLE departments(dept_id INTEGER PRIMARY KEY, dept_name TEXT);
CREATE TABLE employees(emp_id INTEGER PRIMARY KEY, name TEXT, salary INTEGER, dept_id INTEGER, manager_id INTEGER);
CREATE TABLE projects(proj_id INTEGER PRIMARY KEY, proj_name TEXT, dept_id INTEGER);
CREATE TABLE former_employees(emp_id INTEGER, name TEXT, dept_id INTEGER);
INSERT INTO departments VALUES (10,'HR'),(20,'IT'),(30,'Sales'),(40,'Legal');
INSERT INTO employees VALUES
 (1,'Ahmed',9000,20,NULL),(2,'Sara',7000,20,1),(3,'Omar',5000,30,1),
 (4,'Laila',6000,10,1),(5,'Khaled',4500,30,3),(6,'Nora',5200,NULL,1);
INSERT INTO projects VALUES (1,'ERP',20),(2,'CRM',30),(3,'Payroll',10);
INSERT INTO former_employees VALUES (3,'Omar',30),(7,'Yousef',20),(8,'Huda',10);
