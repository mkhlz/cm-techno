
import React, { useState } from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { Search, Filter, BookOpen } from 'lucide-react';
import { Button } from '@/components/ui/button';
import CourseCard from '@/components/CourseCard';
import { courses } from '@/data/courses';

function ITCoursesPage({ onOpenEnquiry }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [sortBy, setSortBy] = useState('newest');

  const categories = ['All', ...new Set(courses.map(c => c.category))];

  const filteredCourses = courses
    .filter(course => {
      const matchesSearch = course.title.toLowerCase().includes(searchTerm.toLowerCase()) || 
                            course.description.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesCategory = selectedCategory === 'All' || course.category === selectedCategory;
      return matchesSearch && matchesCategory;
    })
    .sort((a, b) => {
      if (sortBy === 'name') return a.title.localeCompare(b.title);
      return 0; // Default order
    });

  return (
    <>
      <Helmet>
        <title>All IT Courses - Programming, Data Science, Design | CM Techno Solution</title>
        <meta name="description" content="Explore our wide range of IT courses in Mumbai including Python, Data Science, Digital Marketing, Graphic Design, and more. 100% Practical Training." />
      </Helmet>

      {/* Hero Section */}
      <section className="relative pt-40 pb-16 bg-gradient-to-br from-blue-900 via-blue-800 to-blue-900 text-white overflow-hidden">
        <div className="absolute inset-0 bg-black/20 z-0"></div>
        <div className="container mx-auto px-4 relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
            className="max-w-4xl mx-auto"
          >
            <div className="inline-block px-4 py-2 bg-white/10 backdrop-blur-sm rounded-full border border-white/20 mb-6">
              <span className="text-sm font-medium">🎓 18+ Professional Courses</span>
            </div>
            
            <h1 className="text-4xl md:text-6xl font-bold mb-6 leading-tight">
              Upgrade Your Skills <br />
              <span className="text-blue-300">Boost Your Career</span>
            </h1>
            
            <p className="text-xl text-blue-100 mb-8 max-w-2xl mx-auto">
              Industry-relevant courses designed by experts. Join thousands of successful students who transformed their careers with us.
            </p>
          </motion.div>
        </div>
      </section>

      {/* Filter & Search Section */}
      <section className="py-10 bg-gray-50 border-b border-gray-200 sticky top-16 z-30 shadow-sm">
        <div className="container mx-auto px-4">
          <div className="flex flex-col md:flex-row gap-4 items-center justify-between">
            
            {/* Search */}
            <div className="relative w-full md:w-96">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-5 h-5" />
              <input 
                type="text" 
                placeholder="Search courses..." 
                className="w-full pl-10 pr-4 py-3 rounded-xl border border-gray-200 focus:ring-2 focus:ring-blue-500 outline-none transition-shadow"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>

            {/* Categories - Desktop */}
            <div className="hidden md:flex gap-2 overflow-x-auto pb-1 md:pb-0">
              {categories.map(cat => (
                <button
                  key={cat}
                  onClick={() => setSelectedCategory(cat)}
                  className={`px-4 py-2 rounded-lg text-sm font-medium transition-all whitespace-nowrap ${
                    selectedCategory === cat 
                      ? 'bg-blue-600 text-white shadow-md' 
                      : 'bg-white text-gray-600 hover:bg-gray-100 border border-gray-200'
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>

            {/* Sort - Mobile/Desktop */}
            <select 
              className="md:hidden w-full p-3 rounded-xl border border-gray-200 bg-white"
              value={selectedCategory}
              onChange={(e) => setSelectedCategory(e.target.value)}
            >
              {categories.map(cat => (
                <option key={cat} value={cat}>{cat}</option>
              ))}
            </select>
          </div>
        </div>
      </section>

      {/* Course Grid */}
      <section className="py-16 bg-white min-h-screen">
        <div className="container mx-auto px-4">
          {filteredCourses.length > 0 ? (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-8">
              {filteredCourses.map((course) => (
                <CourseCard key={course.id} course={course} />
              ))}
            </div>
          ) : (
            <div className="text-center py-20">
              <div className="bg-gray-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-4">
                <Search className="w-10 h-10 text-gray-400" />
              </div>
              <h3 className="text-xl font-bold text-gray-900 mb-2">No courses found</h3>
              <p className="text-gray-500">Try adjusting your search or filter to find what you're looking for.</p>
              <Button 
                variant="outline" 
                className="mt-4"
                onClick={() => {setSearchTerm(''); setSelectedCategory('All');}}
              >
                Clear Filters
              </Button>
            </div>
          )}
        </div>
      </section>

      {/* Bottom CTA */}
      <section className="py-16 bg-gradient-to-r from-blue-900 to-blue-800 text-white text-center">
        <div className="container mx-auto px-4">
          <h2 className="text-3xl font-bold mb-4">Can't decide which course to choose?</h2>
          <p className="text-blue-100 mb-8 max-w-xl mx-auto">
            Our career counselors can help you select the right path based on your background and goals.
          </p>
          <Button 
            size="lg" 
            className="bg-white text-blue-900 hover:bg-blue-50 font-bold px-8 shadow-xl"
            onClick={onOpenEnquiry}
          >
            Get Free Career Counseling
          </Button>
        </div>
      </section>
    </>
  );
}

export default ITCoursesPage;
