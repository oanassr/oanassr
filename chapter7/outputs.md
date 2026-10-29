# Verified outputs (SQLite 3.45.1)

## 01 INNER JOIN: link employees to their department names
```sql
SELECT e.name, d.dept_name FROM employees e INNER JOIN departments d ON e.dept_id = d.dept_id;
```
| name | dept_name |
|---|---|
| Ahmed | IT |
| Sara | IT |
| Omar | Sales |
| Laila | HR |
| Khaled | Sales |

## 02 LEFT JOIN: keep ALL employees, even without a department
```sql
SELECT e.name, d.dept_name FROM employees e LEFT JOIN departments d ON e.dept_id = d.dept_id;
```
| name | dept_name |
|---|---|
| Ahmed | IT |
| Sara | IT |
| Omar | Sales |
| Laila | HR |
| Khaled | Sales |
| Nora | NULL |

## 03 RIGHT JOIN: keep ALL departments, even empty ones
```sql
SELECT e.name, d.dept_name FROM employees e RIGHT JOIN departments d ON e.dept_id = d.dept_id;
```
| name | dept_name |
|---|---|
| Ahmed | IT |
| Sara | IT |
| Omar | Sales |
| Laila | HR |
| Khaled | Sales |
| NULL | Legal |

## 04 FULL OUTER JOIN: everything from both sides
```sql
SELECT e.name, d.dept_name FROM employees e FULL OUTER JOIN departments d ON e.dept_id = d.dept_id;
```
| name | dept_name |
|---|---|
| Ahmed | IT |
| Sara | IT |
| Omar | Sales |
| Laila | HR |
| Khaled | Sales |
| Nora | NULL |
| NULL | Legal |

## 05 Anti-join: departments with no employees
```sql
SELECT d.dept_name FROM departments d LEFT JOIN employees e ON e.dept_id = d.dept_id WHERE e.emp_id IS NULL;
```
| dept_name |
|---|
| Legal |

## 06 SELF JOIN: employee with manager
```sql
SELECT e.name AS employee, m.name AS manager FROM employees e LEFT JOIN employees m ON e.manager_id = m.emp_id;
```
| employee | manager |
|---|---|
| Ahmed | NULL |
| Sara | Ahmed |
| Omar | Ahmed |
| Laila | Ahmed |
| Khaled | Omar |
| Nora | Ahmed |

## 07 CROSS JOIN: all combinations (departments x projects count)
```sql
SELECT COUNT(*) AS combos FROM departments CROSS JOIN projects;
```
| combos |
|---|
| 12 |

## 08 Scalar subquery: employees earning above the average
```sql
SELECT name, salary FROM employees WHERE salary > (SELECT AVG(salary) FROM employees);
```
| name | salary |
|---|---|
| Ahmed | 9000 |
| Sara | 7000 |

## 09 IN subquery: employees in departments that have projects
```sql
SELECT name FROM employees WHERE dept_id IN (SELECT dept_id FROM projects);
```
| name |
|---|
| Ahmed |
| Sara |
| Omar |
| Laila |
| Khaled |

## 10 EXISTS: departments having at least one employee
```sql
SELECT dept_name FROM departments d WHERE EXISTS (SELECT 1 FROM employees e WHERE e.dept_id = d.dept_id);
```
| dept_name |
|---|
| HR |
| IT |
| Sales |

## 11 Correlated subquery: top earner in each department
```sql
SELECT name, dept_id, salary FROM employees e WHERE salary = (SELECT MAX(salary) FROM employees WHERE dept_id = e.dept_id);
```
| name | dept_id | salary |
|---|---|---|
| Ahmed | 20 | 9000 |
| Omar | 30 | 5000 |
| Laila | 10 | 6000 |

## 12 UNION: all distinct names, current and former
```sql
SELECT name FROM employees UNION SELECT name FROM former_employees;
```
| name |
|---|
| Ahmed |
| Huda |
| Khaled |
| Laila |
| Nora |
| Omar |
| Sara |
| Yousef |

## 13 UNION ALL: keep duplicates
```sql
SELECT name FROM employees UNION ALL SELECT name FROM former_employees;
```
| name |
|---|
| Ahmed |
| Sara |
| Omar |
| Laila |
| Khaled |
| Nora |
| Omar |
| Yousef |
| Huda |

## 14 INTERSECT: people who are in both lists
```sql
SELECT name FROM employees INTERSECT SELECT name FROM former_employees;
```
| name |
|---|
| Omar |

## 15 EXCEPT: current employees who are not former employees
```sql
SELECT name FROM employees EXCEPT SELECT name FROM former_employees;
```
| name |
|---|
| Ahmed |
| Khaled |
| Laila |
| Nora |
| Sara |
