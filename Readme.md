## About this tool:
While preparing for AWS Certification I wanted to do some sample tests before appearing for the exam. I did not find any free online exams that will help me prepare for the exams and paying $20/- for each practice test was expensive. So I created an s/s with questions and answers that I lifted from “AWS Certified Solutions Architect Official Study Guide”.  My excel files tab “Questions” looks as follows:

 Question number |   Actual Question 
--- | ---
**1**| 1. Which of the following describes a physical location around the world where AWS clusters data centers?
`Empty cell` | A) Endpoint
`Empty cell` | B) Collection
`Empty cell` | C) Fleet
`Empty cell` | **>>>D) Region**
**2**| 2. Each AWS region is composed of two or more locations that offer organizations the ability to operate production systems that are more highly available, fault tolerant, and scalable than would be possible using a single data center. What are these locations called?
`Empty cell` | **>>>A) Availability Zones**
`Empty cell` | B) Replication areas
`Empty cell` | C) Geographic districts
`Empty cell` | D) Computer centers

Now that I have around 225 questions and answers I wanted to simulate the actual test. I use Drill-2 (https://gronostajo.github.io/drill2/) for test simulation. I don’t think it’ll be fair for me to share my s/s as I believe it will be a copyright violation. 

The questions and answers needed to be formatted so that Drill-2 can use them. As it one-off effort I used excel and regular expression to format the questions so that Drill-2 will be happy.

Manual processing to create my excel file with questions:

..*Copied all questions at the end of each chapter into one large text file.

..*Used regular expressions to combine multi lines questions (questions with cr/lf)  to single line.

..* Used regular expressions to format the possible answers to a layout that Drill-2 liked.
...*	(A. Answer => A) Answer)
..* Manually marked the correct answer.
...* (note the >>> placed before the correct answer)
..* Using Excel to organized lines of questions and answers to sets of questions and answers as shown in the above table.
..* My excel file has 231 questions.

If only Drill-2 could all questions and selected random 60 questions then this tool would be redundant. Well, life is not easy and we need to address it! 
