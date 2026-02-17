
import React from 'react';
import { Link } from 'react-router-dom';
import { motion } from 'framer-motion';
import { ArrowRight, CheckCircle } from 'lucide-react';
import { Button } from '@/components/ui/button';

function CourseCard({ course }) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      className="bg-white rounded-2xl shadow-lg hover:shadow-2xl transition-all duration-300 hover:-translate-y-1 flex flex-col h-full overflow-hidden border border-gray-100"
    >
      <div className="relative h-48 overflow-hidden">
        <img
          src={course.image}
          alt={course.title}
          className="w-full h-full object-cover transition-transform duration-500 hover:scale-110"
        />
        <div className="absolute top-4 right-4 bg-white/90 backdrop-blur-sm px-3 py-1 rounded-full text-xs font-bold text-blue-900 shadow-sm">
          {course.duration}
        </div>
      </div>
      
      <div className="p-6 flex-col flex flex-grow">
        <div className="flex items-center mb-3">
          <div className="p-2 bg-blue-50 rounded-lg mr-3">
            <course.icon className="w-5 h-5 text-blue-600" />
          </div>
          <span className="text-xs font-semibold text-blue-600 uppercase tracking-wider bg-blue-50 px-2 py-1 rounded-md">
            {course.category}
          </span>
        </div>

        <h3 className="text-xl font-bold text-gray-900 mb-2 line-clamp-1">{course.title}</h3>
        <p className="text-gray-600 text-sm mb-4 line-clamp-2 flex-grow">{course.description}</p>

        <div className="space-y-2 mb-6">
          {course.highlights.slice(0, 2).map((highlight, index) => (
            <div key={index} className="flex items-center text-xs text-gray-500">
              <CheckCircle className="w-3 h-3 text-green-500 mr-2" />
              {highlight}
            </div>
          ))}
        </div>

        <Link to={`/courses/${course.id}`} className="mt-auto">
          <Button className="w-full bg-blue-50 hover:bg-blue-100 text-blue-700 hover:text-blue-900 border border-blue-200 transition-colors group">
            View Details
            <ArrowRight className="w-4 h-4 ml-2 group-hover:translate-x-1 transition-transform" />
          </Button>
        </Link>
      </div>
    </motion.div>
  );
}

export default CourseCard;
