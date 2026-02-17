
import { 
  Code, BarChart3, TrendingUp, Cpu, Video, 
  Calculator, FileSpreadsheet, Database, Terminal, 
  Layers, Monitor, School
} from 'lucide-react';

export const courses = [
  {
    id: 'data-analytics',
    title: 'Data Analytics Course',
    description: 'Master data analysis tools and techniques to drive business decisions.',
    image: 'https://images.unsplash.com/photo-1516383274235-5f42d6c6426d',
    icon: BarChart3,
    category: 'Data Science',
    duration: '3 Months',
    tools: ['Power BI', 'Excel', 'SQL', 'Python', 'Tableau'],
    overview: 'Become a data expert with our comprehensive Data Analytics course. Learn to collect, process, and analyze data to solve business problems.',
    highlights: ['Live Projects', 'Industry Certification', 'Job Assistance'],
    curriculum: [
      'Introduction to Data Analytics',
      'Advanced Excel for Data Analysis',
      'SQL for Data Science',
      'Data Visualization with Power BI & Tableau',
      'Python for Data Analysis'
    ]
  },
  {
    id: 'digital-marketing',
    title: 'Digital Marketing Course',
    description: 'Complete digital marketing training from SEO to Social Media.',
    image: 'https://images.unsplash.com/photo-1676276375914-cc102df16fa9',
    icon: TrendingUp,
    category: 'Marketing',
    duration: '3 Months',
    tools: ['Google Ads', 'Facebook Ads', 'SEO', 'Analytics', 'Canva'],
    overview: 'Learn to dominate the digital landscape. This course covers SEO, SEM, SMM, Content Marketing, and more.',
    highlights: ['Live Campaign Management', 'Google Certifications', 'Freelancing Training'],
    curriculum: [
      'Digital Marketing Fundamentals',
      'Website Planning & Creation',
      'Search Engine Optimization (SEO)',
      'Social Media Marketing (SMM)',
      'Google Analytics & Reporting'
    ]
  },
  {
    id: 'stock-market',
    title: 'Stock Market Course',
    description: 'Learn technical analysis and trading strategies for the stock market.',
    image: 'https://images.unsplash.com/photo-1500401519266-0b71b29a05e0',
    icon: TrendingUp,
    category: 'Finance',
    duration: '2 Months',
    tools: ['TradingView', 'Screeners', 'Technical Indicators'],
    overview: 'Understand the stock market from basics to advanced technical analysis.',
    highlights: ['Live Trading Sessions', 'Portfolio Management', 'Risk Analysis'],
    curriculum: [
      'Basics of Stock Market',
      'Fundamental Analysis',
      'Technical Analysis',
      'Future & Options',
      'Trading Psychology'
    ]
  },
  {
    id: 'machine-learning',
    title: 'Machine Learning',
    description: 'Build intelligent systems using ML algorithms and Python.',
    image: 'https://images.unsplash.com/photo-1678995635432-d9e89c7a8fc5',
    icon: Cpu,
    category: 'Data Science',
    duration: '4 Months',
    tools: ['Python', 'TensorFlow', 'Scikit-learn', 'Pandas'],
    overview: 'Dive into the world of AI with our Machine Learning course.',
    highlights: ['Real-world Datasets', 'Neural Networks', 'AI Projects'],
    curriculum: [
      'Python for ML',
      'Supervised Learning',
      'Unsupervised Learning',
      'Deep Learning Basics',
      'Model Deployment'
    ]
  },
  {
    id: 'artificial-intelligence',
    title: 'Artificial Intelligence',
    description: 'Advanced AI training covering NLP, Computer Vision, and more.',
    image: 'https://images.unsplash.com/photo-1549925245-f20a1bac6454',
    icon: Cpu,
    category: 'Data Science',
    duration: '6 Months',
    tools: ['Python', 'PyTorch', 'OpenCV', 'Keras'],
    overview: 'Master the technologies shaping the future.',
    highlights: ['Capstone Project', 'Research Papers', 'Industry Expert Sessions'],
    curriculum: [
      'AI Fundamentals',
      'Natural Language Processing',
      'Computer Vision',
      'Reinforcement Learning',
      'Generative AI'
    ]
  },
  {
    id: 'video-editing',
    title: 'Video Editing Course',
    description: 'Professional video editing using Premiere Pro and After Effects.',
    image: 'https://images.unsplash.com/photo-1696389500310-cd6d247cb609',
    icon: Video,
    category: 'Design',
    duration: '3 Months',
    tools: ['Premiere Pro', 'After Effects', 'Photoshop', 'Audition'],
    overview: 'Learn the art of storytelling through video.',
    highlights: ['Portfolio Building', 'Cinematic Techniques', 'Motion Graphics'],
    curriculum: [
      'Video Editing Basics',
      'Color Grading',
      'Audio Mixing',
      'Visual Effects',
      'Rendering & Exporting'
    ]
  },
  {
    id: 'tally-gst',
    title: 'Tally & GST Course',
    description: 'Complete accounting training with Tally Prime and GST filing.',
    image: 'https://images.unsplash.com/photo-1563198804-b144dfc1661c',
    icon: Calculator,
    category: 'Finance',
    duration: '2 Months',
    tools: ['Tally Prime', 'GST Portal', 'Excel'],
    overview: 'Become a professional accountant with practical Tally training.',
    highlights: ['Live GST Filing', 'Payroll Management', 'Taxation'],
    curriculum: [
      'Accounting Principles',
      'Tally Prime Interface',
      'Inventory Management',
      'GST Compliance',
      'Finalization of Accounts'
    ]
  },
  {
    id: 'advance-excel',
    title: 'Advance Excel Course',
    description: 'Master Excel formulas, pivot tables, macros and VBA.',
    image: 'https://images.unsplash.com/photo-1529078155058-5d716f45d604',
    icon: FileSpreadsheet,
    category: 'Data Science',
    duration: '1.5 Months',
    tools: ['Excel', 'VBA', 'Macros'],
    overview: 'Boost your productivity with Advanced Excel skills.',
    highlights: ['Corporate Reporting', 'Automation', 'Dashboard Creation'],
    curriculum: [
      'Advanced Formulas',
      'Data Validation',
      'Pivot Tables & Charts',
      'Macros & VBA',
      'Power Query'
    ]
  },
  {
    id: 'c-cpp',
    title: 'C & C++ Programming',
    description: 'Strong foundation in programming logic and OOP concepts.',
    image: 'https://images.unsplash.com/photo-1555121638-bb997817a76d',
    icon: Terminal,
    category: 'Programming',
    duration: '2 Months',
    tools: ['VS Code', 'Dev C++', 'Compilers'],
    overview: 'Start your coding journey with the mother of all languages.',
    highlights: ['Logic Building', 'System Programming', 'Game Basics'],
    curriculum: [
      'C Language Basics',
      'Pointers & Memory Management',
      'C++ OOP Concepts',
      'File Handling',
      'STL Framework'
    ]
  },
  {
    id: 'dsa',
    title: 'Data Structures & Algo',
    description: 'Master DSA for cracking technical interviews at top tech companies.',
    image: 'https://images.unsplash.com/photo-1699190375905-3cac33bbdbb1',
    icon: Code,
    category: 'Programming',
    duration: '3 Months',
    tools: ['Java/C++', 'LeetCode', 'HackerRank'],
    overview: 'The essential course for any serious software engineer.',
    highlights: ['Interview Prep', 'Competitive Programming', 'Problem Solving'],
    curriculum: [
      'Arrays & Strings',
      'Linked Lists & Stacks',
      'Trees & Graphs',
      'Dynamic Programming',
      'Complexity Analysis'
    ]
  },
  {
    id: 'sql',
    title: 'SQL Database Course',
    description: 'Learn database design, querying, and management.',
    image: 'https://images.unsplash.com/photo-1627398242454-45a1465c2479',
    icon: Database,
    category: 'Data Science',
    duration: '1.5 Months',
    tools: ['MySQL', 'PostgreSQL', 'WorkBench'],
    overview: 'Data is the new oil. Learn to manage it efficiently.',
    highlights: ['Database Design', 'Complex Queries', 'Performance Tuning'],
    curriculum: [
      'RDBMS Concepts',
      'DDL, DML, DCL Commands',
      'Joins & Subqueries',
      'Stored Procedures',
      'Normalization'
    ]
  },
  {
    id: 'java-development',
    title: 'Java Development',
    description: 'Core and Advanced Java with Spring Boot framework.',
    image: 'https://images.unsplash.com/photo-1444300703094-ae12949acab9',
    icon: Code,
    category: 'Programming',
    duration: '4 Months',
    tools: ['Java', 'Spring Boot', 'Hibernate', 'Maven'],
    overview: 'Build robust enterprise applications with Java.',
    highlights: ['Microservices', 'REST APIs', 'Full Stack Project'],
    curriculum: [
      'Core Java',
      'Advanced Java (J2EE)',
      'Spring Framework',
      'Hibernate ORM',
      'Building APIs'
    ]
  },
  {
    id: 'python-development',
    title: 'Python Development',
    description: 'Python for web development, automation, and scripting.',
    image: 'https://images.unsplash.com/photo-1670681423906-0ee3ca2da3ae',
    icon: Code,
    category: 'Programming',
    duration: '3 Months',
    tools: ['Python', 'Django', 'Flask', 'Selenium'],
    overview: 'The most versatile language for modern development.',
    highlights: ['Web Scraping', 'Automation Scripts', 'Backend Dev'],
    curriculum: [
      'Python Syntax & Semantics',
      'OOP in Python',
      'Django Web Framework',
      'REST APIs with DRF',
      'Deployment'
    ]
  },
  {
    id: 'algorithmic-trading',
    title: 'Algorithmic Trading',
    description: 'Automate your trading strategies using Python and APIs.',
    image: 'https://images.unsplash.com/photo-1620266757065-5814239881fd',
    icon: TrendingUp,
    category: 'Finance',
    duration: '3 Months',
    tools: ['Python', 'Broker APIs', 'Backtesting Libraries'],
    overview: 'Combine finance and coding to build trading bots.',
    highlights: ['Live Bot Deployment', 'Strategy Backtesting', 'API Integration'],
    curriculum: [
      'Python for Finance',
      'Connecting to Broker APIs',
      'Building Strategies',
      'Backtesting Frameworks',
      'Risk Management Code'
    ]
  },
  {
    id: 'graphic-designing',
    title: 'Graphic Designing',
    description: 'Master Photoshop, Illustrator, and CorelDraw.',
    image: 'https://images.unsplash.com/photo-1495224814653-94f36c0a31ea',
    icon: Layers,
    category: 'Design',
    duration: '3 Months',
    tools: ['Photoshop', 'Illustrator', 'CorelDraw', 'InDesign'],
    overview: 'Unleash your creativity with professional design tools.',
    highlights: ['Logo Design', 'Branding Kit', 'Social Media Creatives'],
    curriculum: [
      'Design Principles',
      'Image Editing',
      'Vector Graphics',
      'Print Design',
      'Portfolio Creation'
    ]
  },
  {
    id: 'it-engineering',
    title: 'IT Engineering Classes',
    description: 'Support for BSc IT, CS, and Engineering subjects.',
    image: 'https://images.unsplash.com/photo-1521939708078-d6498bb62747',
    icon: Monitor,
    category: 'Academic',
    duration: 'Semester-wise',
    tools: ['Syllabus Specific', 'Lab Work', 'Projects'],
    overview: 'Expert coaching for engineering and degree students.',
    highlights: ['Exam Oriented', 'Practical Labs', 'Project Guidance'],
    curriculum: [
      'Subject Specific Coaching',
      'Practical Assignments',
      'Viva Preparation',
      'Final Year Projects',
      'Doubt Solving'
    ]
  },
  {
    id: 'coding-classes',
    title: 'Coding Classes',
    description: 'General coding foundation for beginners and enthusiasts.',
    image: 'https://images.unsplash.com/photo-1507146815454-9faa99d579aa',
    icon: Code,
    category: 'Programming',
    duration: 'Flexible',
    tools: ['Scratch', 'Logic', 'Basic Syntax'],
    overview: 'Start your programming journey here.',
    highlights: ['Logic Building', 'Small Projects', 'Fun Learning'],
    curriculum: [
      'Introduction to Computers',
      'Programming Logic',
      'Flowcharts & Algorithms',
      'Basic Coding Concepts',
      'First Application'
    ]
  },
  {
    id: 'programming-classes-students',
    title: 'Programming for Students',
    description: 'Specialized coding curriculum for school and college students.',
    image: 'https://images.unsplash.com/photo-1701701046353-89f1a671c24b',
    icon: School,
    category: 'Academic',
    duration: 'Yearly',
    tools: ['School Syllabus', 'Java', 'Python'],
    overview: 'Building the next generation of coders.',
    highlights: ['ICSE/CBSE Syllabus', 'Hands-on Coding', 'Future Ready'],
    curriculum: [
      'School Curriculum Coverage',
      'Advanced Topics',
      'Competitive Coding Basics',
      'Project Work',
      'Certification'
    ]
  }
];
