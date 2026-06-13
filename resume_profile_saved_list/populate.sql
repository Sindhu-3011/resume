-- Populate resume_profiles database with sample data
-- Run this in DBeaver after connecting to resume_profiles

-- Clear existing data (optional)
TRUNCATE resume_skill, resume RESTART IDENTITY CASCADE;

-- Insert sample resumes
INSERT INTO resume (
    full_name, title, email, phone, linkedin, location,
    summary, skills, experience, education, certifications, projects, slug, resume_file
) VALUES
(
    'Sindhu Sundaramoorthy',
    'Validation Engineer',
    'sindhusundaramoorthy30@gmail.com',
    '+91 7358374455',
    'https://www.linkedin.com/in/sindhu-sundaramoorthy',
    'Chennai, India 603103',
    'Validation Engineer with expertise in quality management and compliance within pharmaceutical and IT sectors. Proven ability to develop validation strategies, prepare test cases, and execute protocols aligned with GxP and 21 CFR Part 11 standards.',
    'Validation Testing, GxP Compliance, 21 CFR Part 11, ALM, Test Case Preparation, URS/FRS Review, OQ/PQ Protocols, Manual Testing, Quality Management, Cross-functional Collaboration',
    'VALIDATION ENGINEER | Vaisesika Consulting Pvt Ltd, Bangalore | 11/2021 - Current
- Conducted comprehensive test case preparation and execution, ensuring quality management across projects.
- Collaborated with IT and business teams to finalize validation strategies for GxP projects.
- Reviewed User Requirements Specifications (URS) and Functional Requirements Specifications (FRS) for accuracy.
- Developed assessment checklists for 21 CFR Part 11 and GxP compliance.

PROGRAMMER ANALYST | Cognizant Technology Solutions, Chennai | 06/2018 - 11/2021
- Requirement gathering and analysis.
- Developed Validation Test Plans in accordance with regulatory requirements and SOPs.
- Developed and executed validation protocols (OQ, PQ).
- Conducted manual testing across web, mobile and desktop applications.',
    'B.E. Computer Science | Anna University, Chennai | 2018',
    NULL,
    NULL,
    'sindhu-sundaramoorthy',
    NULL
),
(
    'Arun Sivakumar',
    'Software Engineer',
    'arun.sivakumar@example.com',
    '+91 9876543210',
    'https://www.linkedin.com/in/arun-sivakumar',
    'Bangalore, India',
    'Experienced Software Engineer with 5+ years of expertise in full-stack development, cloud infrastructure, and agile methodologies. Passionate about building scalable and maintainable software solutions.',
    'Python, Java, React, Node.js, PostgreSQL, MySQL, AWS, Docker, Kubernetes, Git, REST APIs, Microservices',
    'SOFTWARE ENGINEER | TCS, Bangalore | 01/2020 - Current
- Developed and maintained microservices architecture using Java Spring Boot.
- Built RESTful APIs consumed by web and mobile clients.
- Deployed applications on AWS using ECS and RDS.
- Reduced deployment time by 40% through CI/CD pipeline improvements.

JUNIOR DEVELOPER | Infosys, Chennai | 06/2018 - 12/2019
- Developed frontend components using React.js.
- Wrote unit and integration tests achieving 85% code coverage.
- Participated in daily standups and sprint planning.',
    'B.Tech Information Technology | VIT University | 2018',
    'AWS Certified Developer - Associate | 2022
Oracle Certified Java Programmer | 2021',
    'E-Commerce Platform | Built scalable product catalog with React + Node.js + PostgreSQL | 2023
Inventory Management System | Real-time stock tracking with WebSocket integration | 2022',
    'arun-sivakumar',
    NULL
),
(
    'Satheesh Kumar',
    'Data Analyst',
    'satheesh.kumar@example.com',
    '+91 9123456780',
    'https://www.linkedin.com/in/satheesh-kumar',
    'Hyderabad, India',
    'Data Analyst with 3 years of experience in transforming raw data into actionable insights. Skilled in SQL, Python, and visualization tools. Strong background in statistical analysis and business intelligence.',
    'SQL, Python, Pandas, NumPy, Tableau, Power BI, Excel, Machine Learning, Data Visualization, ETL, Google Analytics',
    'DATA ANALYST | Wipro, Hyderabad | 03/2021 - Current
- Designed and maintained dashboards in Tableau and Power BI for executive reporting.
- Wrote complex SQL queries to extract and transform data from multiple sources.
- Automated monthly reporting processes saving 20+ hours per month.
- Collaborated with stakeholders to define KPIs and data requirements.

BUSINESS ANALYST INTERN | HCL Technologies | 01/2021 - 03/2021
- Assisted in data collection and cleaning for sales analytics project.
- Created Excel-based reports for the marketing team.',
    'M.Sc Statistics | University of Hyderabad | 2020
B.Sc Mathematics | Osmania University | 2018',
    'Google Data Analytics Certificate | 2022
Microsoft Power BI Data Analyst | 2023',
    'Customer Churn Prediction | ML model with 89% accuracy using Python + Scikit-learn | 2023
Sales Dashboard | Interactive Tableau dashboard tracking 15 KPIs across 5 regions | 2022',
    'satheesh-kumar',
    NULL
);

-- Populate the normalised skills table
INSERT INTO resume_skill (resume_id, skill)
SELECT r.id, TRIM(s.skill)
FROM resume r
CROSS JOIN LATERAL unnest(string_to_array(r.skills, ',')) AS s(skill)
WHERE TRIM(s.skill) != '';

-- Set schema version
DELETE FROM schema_version;
INSERT INTO schema_version (v) VALUES (4);

-- Verify
SELECT id, full_name, title, email FROM resume;
SELECT rs.resume_id, r.full_name, rs.skill FROM resume_skill rs JOIN resume r ON r.id = rs.resume_id ORDER BY rs.resume_id, rs.id;
