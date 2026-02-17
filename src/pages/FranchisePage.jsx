
import React, { useState } from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { 
  Building2, TrendingUp, Users, Award, 
  CheckCircle, Briefcase, DollarSign, MapPin, 
  Phone, MessageCircle, Send
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useToast } from '@/components/ui/use-toast';

function FranchisePage() {
  const { toast } = useToast();
  const [formData, setFormData] = useState({
    name: '',
    email: '',
    phone: '',
    location: '',
    investment: '',
    message: ''
  });

  const handleSubmit = (e) => {
    e.preventDefault();
    toast({
      title: "Application Submitted",
      description: "Thank you for your interest! Our franchise team will contact you shortly.",
    });
    setFormData({ name: '', email: '', phone: '', location: '', investment: '', message: '' });
  };

  const benefits = [
    { icon: TrendingUp, title: 'High ROI', desc: 'Proven business model with excellent returns on investment within 12-18 months.' },
    { icon: Award, title: 'Brand Value', desc: 'Leverage our established brand reputation and student trust.' },
    { icon: Users, title: 'Complete Support', desc: 'End-to-end support in recruitment, training, marketing, and operations.' },
    { icon: Briefcase, title: 'Recession Proof', desc: 'Education sector is ever-growing and resilient to market fluctuations.' }
  ];

  return (
    <>
      <Helmet>
        <title>Franchise Opportunities - Start Your Education Business | CM Techno Solution</title>
        <meta name="description" content="Partner with CM Techno Solution. Low investment, high ROI education franchise opportunity in Mumbai. Join the fastest growing IT training network." />
      </Helmet>

      {/* Hero Section */}
      <section className="relative pt-40 pb-24 bg-gradient-to-br from-blue-900 via-blue-800 to-blue-900 text-white overflow-hidden">
        <div className="absolute inset-0 bg-black/50 z-0"></div>
        <img 
          src="https://images.unsplash.com/photo-1695195274506-347ac4594ff5" 
          alt="Business meeting" 
          className="absolute inset-0 w-full h-full object-cover mix-blend-overlay opacity-40 z-0"
        />
        
        <div className="container mx-auto px-4 relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
            className="max-w-4xl mx-auto"
          >
            <span className="inline-block px-4 py-1.5 bg-white/20 backdrop-blur-md border border-white/30 rounded-full text-sm font-semibold mb-6">
              🚀 Join Our Success Story
            </span>
            <h1 className="text-4xl md:text-6xl font-bold mb-6 leading-tight">
              Franchise Opportunities with <br/>
              <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-200 to-white">CM Techno Solution</span>
            </h1>
            <p className="text-xl text-blue-100 mb-10 max-w-2xl mx-auto">
              Partner with Mumbai's fastest-growing IT Training & Digital Marketing brand. 
              Build a profitable business with our proven model.
            </p>
            
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button size="lg" className="bg-red-600 hover:bg-red-700 text-white px-8 h-14 text-lg shadow-xl hover:scale-105 transition-all" onClick={() => document.getElementById('apply-form').scrollIntoView({ behavior: 'smooth' })}>
                Apply for Franchise
              </Button>
              <a href="https://wa.me/918169809775?text=Hi%2C%20I%20am%20interested%20in%20Franchise" target="_blank" rel="noopener noreferrer">
                <Button variant="outline" size="lg" className="border-white text-blue-900 bg-white hover:bg-gray-100 px-8 h-14 text-lg">
                  <MessageCircle className="w-5 h-5 mr-2" />
                  Chat on WhatsApp
                </Button>
              </a>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Benefits Section */}
      <section className="py-20 bg-white">
        <div className="container mx-auto px-4">
          <div className="text-center mb-16">
            <h2 className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">Why Franchise with Us?</h2>
            <p className="text-gray-600 max-w-2xl mx-auto text-lg">We provide everything you need to start, run, and grow a successful training institute.</p>
          </div>

          <div className="grid md:grid-cols-2 lg:grid-cols-4 gap-8">
            {benefits.map((benefit, index) => (
              <motion.div 
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-8 bg-gray-50 rounded-2xl border border-gray-100 hover:shadow-xl transition-shadow text-center group"
              >
                <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mx-auto mb-6 group-hover:scale-110 transition-transform">
                  <benefit.icon className="w-8 h-8" />
                </div>
                <h3 className="text-xl font-bold text-gray-900 mb-3">{benefit.title}</h3>
                <p className="text-gray-600">{benefit.desc}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Requirements Section */}
      <section className="py-20 bg-gradient-to-br from-blue-900 to-blue-800 text-white">
        <div className="container mx-auto px-4">
          <div className="grid md:grid-cols-2 gap-12 items-center">
            <div>
              <h2 className="text-3xl md:text-4xl font-bold mb-6">Franchise Requirements</h2>
              <p className="text-blue-100 mb-8 text-lg">
                We are looking for passionate partners who are committed to quality education and business growth.
              </p>
              
              <div className="space-y-6">
                <div className="flex items-start">
                  <div className="w-10 h-10 bg-white/10 rounded-lg flex items-center justify-center mr-4 flex-shrink-0">
                    <MapPin className="w-5 h-5 text-yellow-400" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold mb-1">Space Required</h3>
                    <p className="text-blue-200">500 - 1000 Sq. ft carpet area in a commercial location.</p>
                  </div>
                </div>
                <div className="flex items-start">
                  <div className="w-10 h-10 bg-white/10 rounded-lg flex items-center justify-center mr-4 flex-shrink-0">
                    <DollarSign className="w-5 h-5 text-green-400" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold mb-1">Investment</h3>
                    <p className="text-blue-200">₹ 5 - 10 Lakhs (Depending on location and size).</p>
                  </div>
                </div>
                <div className="flex items-start">
                  <div className="w-10 h-10 bg-white/10 rounded-lg flex items-center justify-center mr-4 flex-shrink-0">
                    <Users className="w-5 h-5 text-red-400" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold mb-1">Manpower</h3>
                    <p className="text-blue-200">2 Counselors, 2-3 Trainers, 1 Office Assistant.</p>
                  </div>
                </div>
              </div>
            </div>
            
            <div className="bg-white/10 backdrop-blur-sm p-8 rounded-2xl border border-white/20">
              <h3 className="text-2xl font-bold mb-6">What We Provide</h3>
              <ul className="space-y-4">
                {[
                  "Brand License & Recognition",
                  "Staff Recruitment & Training Assistance",
                  "Marketing & Lead Generation Support",
                  "Course Material & Syllabus",
                  "CRM & Management Software",
                  "Placement Support for Students"
                ].map((item, i) => (
                  <li key={i} className="flex items-center text-blue-50">
                    <CheckCircle className="w-5 h-5 text-green-400 mr-3 flex-shrink-0" />
                    {item}
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      </section>

      {/* Application Form */}
      <section id="apply-form" className="py-20 bg-gray-50">
        <div className="container mx-auto px-4">
          <div className="max-w-3xl mx-auto bg-white rounded-3xl shadow-xl overflow-hidden">
            <div className="bg-blue-600 p-8 text-center text-white">
              <h2 className="text-3xl font-bold mb-2">Apply Now</h2>
              <p className="opacity-90">Take the first step towards your entrepreneurial journey</p>
            </div>
            
            <div className="p-8 md:p-12">
              <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">Full Name</label>
                  <input 
                    type="text" 
                    required 
                    value={formData.name}
                    onChange={(e) => setFormData({...formData, name: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="John Doe"
                  />
                </div>
                
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Email Address</label>
                  <input 
                    type="email" 
                    required 
                    value={formData.email}
                    onChange={(e) => setFormData({...formData, email: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="john@example.com"
                  />
                </div>
                
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Phone Number</label>
                  <input 
                    type="tel" 
                    required 
                    value={formData.phone}
                    onChange={(e) => setFormData({...formData, phone: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="+91 98765 43210"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Preferred Location</label>
                  <input 
                    type="text" 
                    required 
                    value={formData.location}
                    onChange={(e) => setFormData({...formData, location: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="City / Area"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Investment Capacity</label>
                  <select 
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    value={formData.investment}
                    onChange={(e) => setFormData({...formData, investment: e.target.value})}
                  >
                    <option value="">Select Range</option>
                    <option value="5-10L">₹ 5 - 10 Lakhs</option>
                    <option value="10-20L">₹ 10 - 20 Lakhs</option>
                    <option value="20L+">Above ₹ 20 Lakhs</option>
                  </select>
                </div>

                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">Message (Optional)</label>
                  <textarea 
                    rows="3" 
                    value={formData.message}
                    onChange={(e) => setFormData({...formData, message: e.target.value})}
                    className="w-full px-4 py-3 rounded-lg bg-gray-50 border border-gray-300 focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="Any specific questions or details..."
                  ></textarea>
                </div>

                <div className="md:col-span-2 pt-4">
                  <Button type="submit" size="lg" className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold h-12 text-lg">
                    Submit Application <Send className="w-5 h-5 ml-2" />
                  </Button>
                </div>
              </form>
            </div>
          </div>
        </div>
      </section>

      {/* Contact CTA */}
      <section className="py-12 bg-white text-center">
        <div className="container mx-auto px-4">
          <p className="text-gray-600 mb-6 text-lg">Have questions? Talk to our franchise manager directly.</p>
          <div className="flex justify-center gap-4">
             <a href="tel:+918169809775">
              <Button size="lg" variant="outline" className="border-green-600 text-green-700 hover:bg-green-50">
                <Phone className="w-5 h-5 mr-2" /> Call Now
              </Button>
            </a>
            <a href="https://wa.me/918169809775?text=Hi%2C%20Enquiry%20for%20Franchise">
              <Button size="lg" className="bg-green-600 hover:bg-green-700 text-white">
                <MessageCircle className="w-5 h-5 mr-2" /> WhatsApp
              </Button>
            </a>
          </div>
        </div>
      </section>
    </>
  );
}

export default FranchisePage;
