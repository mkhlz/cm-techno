
import React from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { 
  Award, Target, BarChart3, MapPin, 
  FileText, DollarSign, CheckCircle, Star,
  Users, Zap, TrendingUp, Shield
} from 'lucide-react';
import { Button } from '@/components/ui/button';

function WhyChooseUsPage({ onOpenEnquiry }) {
  const differentiators = [
    {
      icon: Award,
      title: 'Practical IT Training',
      description: 'Hands-on learning with real projects, not just theory. Our students work on live projects and build portfolios that land them jobs.',
      color: 'blue'
    },
    {
      icon: Target,
      title: 'ROI-Focused Marketing',
      description: 'We focus on results, not vanity metrics. Every campaign is designed to generate measurable business outcomes and revenue.',
      color: 'green'
    },
    {
      icon: TrendingUp,
      title: 'Performance-Based Strategy',
      description: 'Data-driven approach with continuous optimization. We analyze, test, and improve campaigns for maximum performance.',
      color: 'purple'
    },
    {
      icon: MapPin,
      title: 'Local Market Expertise',
      description: 'Deep understanding of Mumbai market dynamics. We know what works for local businesses and real estate in Malad and Mumbai.',
      color: 'red'
    },
    {
      icon: FileText,
      title: 'Transparent Reporting',
      description: 'Complete visibility into campaign performance. Regular reports with detailed analytics and insights on every rupee spent.',
      color: 'orange'
    },
    {
      icon: DollarSign,
      title: 'Affordable Pricing',
      description: 'Competitive pricing without compromising quality. Flexible packages designed for startups to enterprises.',
      color: 'teal'
    }
  ];

  const comparisonFeatures = [
    { feature: 'Practical Training with Live Projects', us: true, others: false },
    { feature: 'Industry-Experienced Trainers', us: true, others: false },
    { feature: 'Job Placement Assistance', us: true, others: false },
    { feature: 'Real Estate Lead Generation Expertise', us: true, others: false },
    { feature: 'Local Business Marketing Specialization', us: true, others: false },
    { feature: 'Transparent ROI Reporting', us: true, others: false },
    { feature: 'Dedicated Account Manager', us: true, others: false },
    { feature: 'Flexible Payment Options', us: true, others: false }
  ];

  const testimonials = [
    {
      name: 'Amit Kumar',
      role: 'Full Stack Developer',
      image: 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=100&h=100&fit=crop',
      text: 'CM Techno Solution changed my career! The practical training approach and live projects helped me get a job within 2 months of course completion.',
      rating: 5
    },
    {
      name: 'Sneha Reddy',
      role: 'Real Estate Developer',
      image: 'https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=100&h=100&fit=crop',
      text: 'Their real estate lead generation service is exceptional. We closed ₹15 Cr worth of deals in just 4 months. Highly recommended!',
      rating: 5
    },
    {
      name: 'Dr. Rajesh Patel',
      role: 'Clinic Owner',
      image: 'https://images.unsplash.com/photo-1500648767791-00dcc994a43e?w=100&h=100&fit=crop',
      text: 'Patient bookings increased by 180% after working with them. Professional team with deep understanding of local market.',
      rating: 5
    },
    {
      name: 'Meera Shah',
      role: 'E-commerce Business Owner',
      image: 'https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=100&h=100&fit=crop',
      text: 'ROI-focused approach is what sets them apart. Our Shopify sales tripled and customer acquisition cost reduced by 40%.',
      rating: 5
    }
  ];

  const stats = [
    { icon: Users, value: '500+', label: 'Students Trained' },
    { icon: Award, value: '100+', label: 'Projects Completed' },
    { icon: TrendingUp, value: '5x', label: 'Average Client ROI' },
    { icon: Star, value: '4.9/5', label: 'Customer Rating' }
  ];

  return (
    <>
      <Helmet>
        <title>Why Choose CM Techno Solution - Best IT Training & Digital Marketing in Mumbai</title>
        <meta 
          name="description" 
          content="Discover why CM Techno Solution is the best choice for IT training and digital marketing in Mumbai. Practical training, ROI-focused marketing, transparent reporting, and affordable pricing." 
        />
      </Helmet>

      {/* Hero Section */}
      <section className="relative pt-32 pb-16 bg-gradient-to-br from-blue-900 via-blue-800 to-blue-900 text-white overflow-hidden">
        <div className="absolute inset-0 opacity-10">
          <div className="absolute inset-0 bg-[url('data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNjAiIGhlaWdodD0iNjAiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PGRlZnM+PHBhdHRlcm4gaWQ9ImdyaWQiIHdpZHRoPSI2MCIgaGVpZ2h0PSI2MCIgcGF0dGVyblVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+PHBhdGggZD0iTSAxMCAwIEwgMCAwIDAgMTAiIGZpbGw9Im5vbmUiIHN0cm9rZT0id2hpdGUiIHN0cm9rZS13aWR0aD0iMSIvPjwvcGF0dGVybj48L2RlZnM+PHJlY3Qgd2lkdGg9IjEwMCUiIGhlaWdodD0iMTAwJSIgZmlsbD0idXJsKCNncmlkKSIvPjwvc3ZnPg==')]"></div>
        </div>

        <div className="container mx-auto px-4 relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
            className="text-center max-w-4xl mx-auto"
          >
            <div className="inline-block px-4 py-2 bg-white/10 backdrop-blur-sm rounded-full border border-white/20 mb-6">
              <span className="text-sm font-medium">⭐ Trusted by 500+ Students & 100+ Businesses</span>
            </div>
            
            <h1 className="text-4xl md:text-5xl lg:text-6xl font-bold mb-6 leading-tight">
              Why Choose <span className="text-blue-300">CM Techno Solution</span>?
            </h1>
            
            <p className="text-xl md:text-2xl text-blue-100 mb-8">
              Excellence in IT Training & Digital Marketing with Proven Results
            </p>
          </motion.div>
        </div>
      </section>

      {/* Stats Section */}
      <section className="py-12 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
            {stats.map((stat, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="text-center p-6 bg-gradient-to-br from-blue-50 to-blue-100 rounded-2xl shadow-lg"
              >
                <stat.icon className="w-12 h-12 mx-auto mb-3 text-blue-600" />
                <div className="text-3xl font-bold text-blue-900 mb-2">{stat.value}</div>
                <div className="text-sm text-gray-700 font-medium">{stat.label}</div>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Key Differentiators */}
      <section className="py-16 bg-gradient-to-br from-gray-50 to-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              What Makes Us Different
            </h2>
            <p className="text-xl text-gray-600 max-w-2xl mx-auto">
              Our unique approach combines practical training with result-oriented marketing
            </p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
            {differentiators.map((item, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-6 bg-white rounded-2xl shadow-lg hover:shadow-2xl transition-all hover:scale-105"
              >
                <div className={`w-16 h-16 bg-${item.color}-100 rounded-2xl flex items-center justify-center mb-4`}>
                  <item.icon className={`w-8 h-8 text-${item.color}-600`} />
                </div>
                <h3 className="text-xl font-bold text-gray-900 mb-3">{item.title}</h3>
                <p className="text-gray-600">{item.description}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Comparison Section */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              CM Techno Solution vs Others
            </h2>
            <p className="text-xl text-gray-600">See how we compare to our competitors</p>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="max-w-4xl mx-auto"
          >
            <div className="bg-white rounded-2xl shadow-xl overflow-hidden">
              <div className="grid grid-cols-3 bg-gradient-to-r from-blue-600 to-blue-800 text-white p-4 font-bold">
                <div className="col-span-1">Features</div>
                <div className="text-center">CM Techno Solution</div>
                <div className="text-center">Others</div>
              </div>
              
              {comparisonFeatures.map((item, index) => (
                <div
                  key={index}
                  className={`grid grid-cols-3 p-4 border-b border-gray-200 ${
                    index % 2 === 0 ? 'bg-gray-50' : 'bg-white'
                  }`}
                >
                  <div className="col-span-1 text-gray-900 font-medium">{item.feature}</div>
                  <div className="text-center">
                    {item.us ? (
                      <CheckCircle className="w-6 h-6 text-green-600 mx-auto" />
                    ) : (
                      <span className="text-gray-400">✕</span>
                    )}
                  </div>
                  <div className="text-center">
                    {item.others ? (
                      <CheckCircle className="w-6 h-6 text-green-600 mx-auto" />
                    ) : (
                      <span className="text-gray-400">✕</span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </motion.div>
        </div>
      </section>

      {/* Featured Image Section */}
      <section className="py-16 bg-gradient-to-r from-blue-600 to-blue-800 text-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center">
            <motion.div
              initial={{ opacity: 0, x: -20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <img
                src="https://images.unsplash.com/flagged/photo-1551135049-83f3419ef05c?w=800&h=600&fit=crop"
                alt="Professional team collaboration at CM Techno Solution office"
                className="rounded-2xl shadow-2xl"
              />
            </motion.div>

            <motion.div
              initial={{ opacity: 0, x: 20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <h2 className="text-4xl font-bold mb-6">
                Your Success is Our Mission
              </h2>
              <p className="text-xl text-blue-100 mb-6">
                We're committed to delivering exceptional results, whether you're looking to build your IT career or grow your business through digital marketing.
              </p>
              <ul className="space-y-4">
                <li className="flex items-center">
                  <Shield className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>10+ years of industry experience</span>
                </li>
                <li className="flex items-center">
                  <Users className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Expert team of trainers and marketers</span>
                </li>
                <li className="flex items-center">
                  <Zap className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Cutting-edge tools and technologies</span>
                </li>
                <li className="flex items-center">
                  <TrendingUp className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Proven track record of success</span>
                </li>
              </ul>
            </motion.div>
          </div>
        </div>
      </section>

      {/* Testimonials Section */}
      <section className="py-16 bg-gradient-to-br from-gray-50 to-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              What Our Clients Say
            </h2>
            <p className="text-xl text-gray-600">Real testimonials from satisfied students and businesses</p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            {testimonials.map((testimonial, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-6 bg-white rounded-2xl shadow-lg hover:shadow-xl transition-shadow"
              >
                <div className="flex items-center mb-4">
                  <img
                    src={testimonial.image}
                    alt={testimonial.name}
                    className="w-12 h-12 rounded-full object-cover mr-3"
                  />
                  <div>
                    <div className="font-bold text-gray-900">{testimonial.name}</div>
                    <div className="text-sm text-gray-600">{testimonial.role}</div>
                  </div>
                </div>
                <div className="flex mb-3">
                  {[...Array(testimonial.rating)].map((_, i) => (
                    <Star key={i} className="w-4 h-4 fill-yellow-400 text-yellow-400" />
                  ))}
                </div>
                <p className="text-gray-700 text-sm">{testimonial.text}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-gradient-to-br from-blue-900 to-blue-800 text-white">
        <div className="container mx-auto px-4 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
          >
            <h2 className="text-4xl font-bold mb-4">Ready to Get Started?</h2>
            <p className="text-xl text-blue-100 mb-8 max-w-2xl mx-auto">
              Join hundreds of successful students and businesses who have transformed their careers and grown their businesses with us
            </p>
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button
                onClick={onOpenEnquiry}
                className="px-8 py-6 bg-white hover:bg-gray-100 text-blue-900 rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl"
              >
                Enquire Now
              </Button>
              <a href="tel:+919876543210">
                <Button className="px-8 py-6 bg-green-600 hover:bg-green-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl">
                  Call Now
                </Button>
              </a>
            </div>
          </motion.div>
        </div>
      </section>
    </>
  );
}

export default WhyChooseUsPage;
