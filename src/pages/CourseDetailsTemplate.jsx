
import React, { useState } from 'react';
import { useParams, Navigate, Link } from 'react-router-dom';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { 
  CheckCircle, Clock, Award, Briefcase, BookOpen, 
  Terminal, Monitor, Users, ArrowRight, Star,
  Phone, MessageCircle, Send
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useToast } from '@/components/ui/use-toast';
import { courses } from '@/data/courses';

function CourseDetailsTemplate({ onOpenEnquiry }) {
  const { slug } = useParams();
  const { toast } = useToast();
  const course = courses.find(c => c.id === slug);

  // Form state
  const [formData, setFormData] = useState({
    name: '',
    phone: '',
    email: ''
  });

  if (!course) {
    return <Navigate to="/it-courses" replace />;
  }

  const relatedCourses = courses
    .filter(c => c.category === course.category && c.id !== course.id)
    .slice(0, 3);

  const handleFormSubmit = (e) => {
    e.preventDefault();
    toast({
      title: "Enquiry Submitted",
      description: `Thanks ${formData.name}, we'll contact you about the ${course.title} soon!`,
    });
    setFormData({ name: '', phone: '', email: '' });
  };

  return (
    <>
      <Helmet>
        <title>{`${course.title} - Certification Course in Mumbai | CM Techno Solution`}</title>
        <meta name="description" content={`Enroll in ${course.title} at CM Techno Solution. ${course.description} 100% Placement Assistance, Live Projects, and Internship.`} />
      </Helmet>

      {/* Sticky Header for Mobile */}
      <div className="fixed top-16 left-0 right-0 z-40 bg-white/95 backdrop-blur-sm border-b border-gray-200 py-2 px-4 md:hidden shadow-sm">
        <div className="flex items-center justify-between">
          <span className="font-bold text-gray-900 text-sm truncate pr-2">{course.title}</span>
          <Button size="sm" onClick={onOpenEnquiry} className="bg-blue-600 text-white text-xs px-3 h-8">
            Enroll Now
          </Button>
        </div>
      </div>

      {/* Hero Section */}
      <section className="relative pt-32 pb-20 bg-gradient-to-br from-blue-900 via-blue-800 to-blue-900 text-white overflow-hidden">
        <div className="absolute inset-0 bg-black/40 z-0"></div>
        <img 
          src={course.image} 
          alt={course.title} 
          className="absolute inset-0 w-full h-full object-cover mix-blend-overlay opacity-30 z-0"
        />
        
        <div className="container mx-auto px-4 relative z-10">
          <div className="flex flex-col md:flex-row gap-12 items-center">
            <motion.div 
              initial={{ opacity: 0, y: 30 }}
              animate={{ opacity: 1, y: 0 }}
              className="flex-1"
            >
              <div className="inline-flex items-center space-x-2 bg-white/10 backdrop-blur-md px-4 py-1.5 rounded-full border border-white/20 mb-6">
                <Award className="w-4 h-4 text-yellow-400" />
                <span className="text-sm font-medium">Certification Course</span>
              </div>
              
              <h1 className="text-4xl md:text-5xl lg:text-6xl font-bold mb-6 leading-tight">
                {course.title}
              </h1>
              
              <p className="text-xl text-blue-100 mb-8 leading-relaxed max-w-2xl">
                {course.overview}
              </p>

              <div className="flex flex-wrap gap-4 mb-8">
                <div className="flex items-center bg-blue-800/50 px-4 py-2 rounded-lg border border-blue-700">
                  <Clock className="w-5 h-5 mr-2 text-blue-300" />
                  <span>{course.duration}</span>
                </div>
                <div className="flex items-center bg-blue-800/50 px-4 py-2 rounded-lg border border-blue-700">
                  <Briefcase className="w-5 h-5 mr-2 text-blue-300" />
                  <span>Job Assistance</span>
                </div>
                <div className="flex items-center bg-blue-800/50 px-4 py-2 rounded-lg border border-blue-700">
                  <Monitor className="w-5 h-5 mr-2 text-blue-300" />
                  <span>Online / Offline</span>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row gap-4">
                <Button onClick={onOpenEnquiry} size="lg" className="bg-red-600 hover:bg-red-700 text-white font-bold text-lg px-8 shadow-xl hover:scale-105 transition-transform">
                  Enroll Now
                </Button>
                <Button variant="outline" size="lg" className="bg-white/10 border-white text-white hover:bg-white hover:text-blue-900 font-bold text-lg px-8 backdrop-blur-sm">
                  Download Syllabus
                </Button>
              </div>
            </motion.div>

            {/* Floating Form Card */}
            <motion.div 
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              transition={{ delay: 0.2 }}
              className="w-full md:w-96 bg-white rounded-2xl p-6 shadow-2xl border border-gray-100 text-gray-900"
            >
              <h3 className="text-xl font-bold mb-2">Book Free Demo Class</h3>
              <p className="text-gray-500 text-sm mb-6">Fill the form to get course details & fee structure.</p>
              
              <form onSubmit={handleFormSubmit} className="space-y-4">
                <div>
                  <input
                    type="text"
                    placeholder="Your Name"
                    required
                    value={formData.name}
                    onChange={(e) => setFormData({...formData, name: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 transition-all outline-none"
                  />
                </div>
                <div>
                  <input
                    type="tel"
                    placeholder="Phone Number"
                    required
                    value={formData.phone}
                    onChange={(e) => setFormData({...formData, phone: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 transition-all outline-none"
                  />
                </div>
                <div>
                  <input
                    type="email"
                    placeholder="Email Address"
                    required
                    value={formData.email}
                    onChange={(e) => setFormData({...formData, email: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 transition-all outline-none"
                  />
                </div>
                <input type="hidden" value={course.title} />
                
                <Button type="submit" className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 rounded-lg shadow-lg hover:shadow-xl transition-all">
                  Get Callback
                </Button>
              </form>
            </motion.div>
          </div>
        </div>
      </section>

      {/* Course Highlights & Overview */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid md:grid-cols-3 gap-12">
            <div className="md:col-span-2">
              <h2 className="text-3xl font-bold text-gray-900 mb-6">Course Overview</h2>
              <p className="text-gray-600 text-lg mb-8 leading-relaxed">
                {course.description} Our {course.title} is designed for both beginners and professionals looking to upgrade their skills. 
                With a practical-first approach, you will work on real-world projects and gain hands-on experience that employers value.
              </p>

              <h3 className="text-2xl font-bold text-gray-900 mb-4">What You Will Learn?</h3>
              <div className="grid sm:grid-cols-2 gap-4 mb-10">
                {course.curriculum.map((item, index) => (
                  <div key={index} className="flex items-center p-3 bg-blue-50 rounded-lg border border-blue-100">
                    <CheckCircle className="w-5 h-5 text-blue-600 mr-3 flex-shrink-0" />
                    <span className="font-medium text-gray-800">{item}</span>
                  </div>
                ))}
              </div>

              <h3 className="text-2xl font-bold text-gray-900 mb-4">Tools & Technologies</h3>
              <div className="flex flex-wrap gap-3 mb-10">
                {course.tools.map((tool, index) => (
                  <span key={index} className="px-4 py-2 bg-gray-100 text-gray-700 rounded-full font-medium border border-gray-200">
                    {tool}
                  </span>
                ))}
              </div>

              {/* Badges Section */}
              <div className="bg-gradient-to-r from-blue-900 to-blue-800 rounded-2xl p-8 text-white mb-10 shadow-xl relative overflow-hidden">
                <div className="absolute top-0 right-0 w-64 h-64 bg-white/5 rounded-full -mr-16 -mt-16 blur-3xl"></div>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-6 relative z-10">
                  <div className="text-center">
                    <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center mx-auto mb-3 backdrop-blur-sm">
                      <Briefcase className="w-6 h-6 text-white" />
                    </div>
                    <div className="font-bold">100% Placement Support</div>
                  </div>
                  <div className="text-center">
                    <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center mx-auto mb-3 backdrop-blur-sm">
                      <Terminal className="w-6 h-6 text-white" />
                    </div>
                    <div className="font-bold">Free Internship</div>
                  </div>
                  <div className="text-center">
                    <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center mx-auto mb-3 backdrop-blur-sm">
                      <Award className="w-6 h-6 text-white" />
                    </div>
                    <div className="font-bold">ISO Certified</div>
                  </div>
                  <div className="text-center">
                    <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center mx-auto mb-3 backdrop-blur-sm">
                      <BookOpen className="w-6 h-6 text-white" />
                    </div>
                    <div className="font-bold">Live Projects</div>
                  </div>
                </div>
              </div>

            </div>

            {/* Sidebar */}
            <div className="md:col-span-1">
              <div className="sticky top-24 space-y-6">
                <div className="bg-gray-50 rounded-2xl p-6 border border-gray-200">
                  <h3 className="font-bold text-lg mb-4">Why Join Us?</h3>
                  <ul className="space-y-3">
                    {course.highlights.map((highlight, idx) => (
                      <li key={idx} className="flex items-start text-sm text-gray-600">
                        <Star className="w-4 h-4 text-yellow-500 mr-2 mt-0.5 fill-yellow-500" />
                        {highlight}
                      </li>
                    ))}
                    <li className="flex items-start text-sm text-gray-600">
                      <Star className="w-4 h-4 text-yellow-500 mr-2 mt-0.5 fill-yellow-500" />
                      Experienced Trainers
                    </li>
                    <li className="flex items-start text-sm text-gray-600">
                      <Star className="w-4 h-4 text-yellow-500 mr-2 mt-0.5 fill-yellow-500" />
                      Small Batch Size
                    </li>
                  </ul>
                </div>

                <div className="bg-blue-50 rounded-2xl p-6 border border-blue-100 text-center">
                  <h3 className="font-bold text-lg text-blue-900 mb-2">Need Help?</h3>
                  <p className="text-blue-700 text-sm mb-4">Talk to our career counselor today.</p>
                  <a href="tel:+918169809775" className="block mb-3">
                    <Button className="w-full bg-blue-600 hover:bg-blue-700 text-white">
                      <Phone className="w-4 h-4 mr-2" />
                      Call Now
                    </Button>
                  </a>
                  <a href="https://wa.me/918169809775?text=Hi%2C%20I%20am%20interested%20in%20a%20course" target="_blank" rel="noopener noreferrer">
                    <Button variant="outline" className="w-full border-green-500 text-green-600 hover:bg-green-50">
                      <MessageCircle className="w-4 h-4 mr-2" />
                      WhatsApp
                    </Button>
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Related Courses */}
      {relatedCourses.length > 0 && (
        <section className="py-16 bg-gray-50 border-t border-gray-200">
          <div className="container mx-auto px-4">
            <h2 className="text-3xl font-bold text-gray-900 mb-8 text-center">Related Courses</h2>
            <div className="grid md:grid-cols-3 gap-6">
              {relatedCourses.map((related) => (
                <Link to={`/courses/${related.id}`} key={related.id} className="block group">
                  <div className="bg-white rounded-xl shadow-sm hover:shadow-lg transition-all overflow-hidden border border-gray-200">
                    <div className="h-40 overflow-hidden">
                      <img src={related.image} alt={related.title} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-500" />
                    </div>
                    <div className="p-4">
                      <h3 className="font-bold text-lg text-gray-900 mb-2">{related.title}</h3>
                      <p className="text-sm text-gray-500 line-clamp-2">{related.description}</p>
                      <div className="mt-4 flex items-center text-blue-600 font-medium text-sm">
                        View Details <ArrowRight className="w-4 h-4 ml-1" />
                      </div>
                    </div>
                  </div>
                </Link>
              ))}
            </div>
          </div>
        </section>
      )}

      {/* Sticky Bottom CTA for Mobile */}
      <div className="fixed bottom-0 left-0 right-0 z-50 bg-white border-t border-gray-200 p-3 md:hidden flex gap-2 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.1)]">
        <a href="tel:+918169809775" className="flex-1">
          <Button variant="outline" className="w-full border-blue-600 text-blue-600">
            <Phone className="w-4 h-4 mr-2" /> Call
          </Button>
        </a>
        <Button onClick={onOpenEnquiry} className="flex-1 bg-red-600 hover:bg-red-700 text-white">
          Enroll Now
        </Button>
      </div>
    </>
  );
}

export default CourseDetailsTemplate;
