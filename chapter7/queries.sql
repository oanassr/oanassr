-- 01 INNER JOIN: link employees to their department names
SELECT e.name, d.dept_name FROM employees e INNER JOIN departments d ON e.dept_id = d.dept_id;
-- 02 LEFT JOIN: keep ALL employees, even without a department
SELECT e.name, d.dept_name FROM employees e LEFT JOIN departments d ON e.dept_id = d.dept_id;
-- 03 RIGHT JOIN: keep ALL departments, even empty ones
SELECT e.name, d.dept_name FROM employees e RIGHT JOIN departments d ON e.dept_id = d.dept_id;
-- 04 FULL OUTER JOIN: everything from both sides
SELECT e.name, d.dept_name FROM employees e FULL OUTER JOIN departments d ON e.dept_id = d.dept_id;
-- 05 Anti-join: departments with no employees
SELECT d.dept_name FROM departments d LEFT JOIN employees e ON e.dept_id = d.dept_id WHERE e.emp_id IS NULL;
-- 06 SELF JOIN: employee with manager
SELECT e.name AS employee, m.name AS manager FROM employees e LEFT JOIN employees m ON e.manager_id = m.emp_id;
-- 07 CROSS JOIN: all combinations (departments x projects count)
SELECT COUNT(*) AS combos FROM departments CROSS JOIN projects;
-- 08 Scalar subquery: employees earning above the average
SELECT name, salary FROM employees WHERE salary > (SELECT AVG(salary) FROM employees);
-- 09 IN subquery: employees in departments that have projects
SELECT name FROM employees WHERE dept_id IN (SELECT dept_id FROM projects);
-- 10 EXISTS: departments having at least one employee
SELECT dept_name FROM departments d WHERE EXISTS (SELECT 1 FROM employees e WHERE e.dept_id = d.dept_id);
-- 11 Correlated subquery: top earner in each department
SELECT name, dept_id, salary FROM employees e WHERE salary = (SELECT MAX(salary) FROM employees WHERE dept_id = e.dept_id);
-- 12 UNION: all distinct names, current and former
SELECT name FROM employees UNION SELECT name FROM former_employees;
-- 13 UNION ALL: keep duplicates
SELECT name FROM employees UNION ALL SELECT name FROM former_employees;
-- 14 INTERSECT: people who are in both lists
SELECT name FROM employees INTERSECT SELECT name FROM former_employees;
-- 15 EXCEPT: current employees who are not former employees
SELECT name FROM employees EXCEPT SELECT name FROM former_employees;
