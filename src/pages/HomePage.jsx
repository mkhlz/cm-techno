
import React from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { Link } from 'react-router-dom';
import { 
  Users, Award, TrendingUp, Code, Target, 
  BarChart, Rocket, CheckCircle, Star, ArrowRight,
  Building2, ShoppingBag, LineChart
} from 'lucide-react';
import { Button } from '@/components/ui/button';

function HomePage({ onOpenEnquiry }) {
  const stats = [
    { icon: Users, label: 'Over 20,000 Students Trained', value: '20,000+' },
    { icon: Award, label: 'Total Courses', value: '15+' },
    { icon: TrendingUp, label: 'Years Experience', value: '23+' },
    { icon: Star, label: '4.9/5 Rating', value: '4.9' }
  ];

  const services = [
    {
      icon: Code,
      title: 'IT Training Courses',
      description: 'Python, Full Stack Development, Data Analysis, Digital Marketing, Stock Market Trading, Tally & more',
      color: 'blue',
      link: '/it-courses'
    },
    {
      icon: Building2,
      title: 'Real Estate Lead Generation',
      description: 'High-intent buyer leads for ready-to-move and under-construction properties',
      color: 'green',
      link: '/digital-marketing-services'
    },
    {
      icon: Target,
      title: 'Local Business Marketing',
      description: 'Targeted campaigns for doctors, gyms, salons, CAs and professional services',
      color: 'purple',
      link: '/digital-marketing-services'
    },
    {
      icon: ShoppingBag,
      title: 'E-commerce Marketing',
      description: 'Shopify marketing, Facebook & Instagram ads, conversion optimization',
      color: 'orange',
      link: '/digital-marketing-services'
    },
    {
      icon: LineChart,
      title: 'Google & Meta Ads',
      description: 'Performance marketing campaigns with measurable ROI and transparent reporting',
      color: 'red',
      link: '/digital-marketing-services'
    },
    {
      icon: Rocket,
      title: 'Website Development',
      description: 'Professional websites optimized for conversions and search engines',
      color: 'teal',
      link: '/digital-marketing-services'
    }
  ];

  const testimonials = [
    {
      name: 'Priya Sharma',
      role: 'Python Developer',
      image: 'https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=100&h=100&fit=crop',
      text: 'The Python course was excellent! Practical approach with real projects helped me land my first job as a developer.',
      rating: 5
    },
    {
      name: 'Rahul Mehta',
      role: 'Real Estate Agent',
      image: 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=100&h=100&fit=crop',
      text: 'Their real estate lead generation service is outstanding. We closed 15 deals in 3 months with quality leads.',
      rating: 5
    },
    {
      name: 'Dr. Anjali Desai',
      role: 'Dental Clinic Owner',
      image: 'https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=100&h=100&fit=crop',
      text: 'Digital marketing services increased our patient bookings by 200%. Highly professional and result-oriented team.',
      rating: 5
    },
    {
      name: 'Vikram Patel',
      role: 'E-commerce Business Owner',
      image: 'https://images.unsplash.com/photo-1500648767791-00dcc994a43e?w=100&h=100&fit=crop',
      text: 'Their Facebook ads strategy transformed our Shopify store. Sales increased by 300% in just 2 months!',
      rating: 5
    }
  ];

  const caseStudies = [
    {
      title: 'Real Estate Project - 50 Cr Sales',
      description: 'Generated ₹50 Cr in property sales through targeted Facebook and Google ads campaign',
      metrics: '₹50 Cr Revenue',
      image: 'https://images.unsplash.com/photo-1698316738298-7f92b28225e4?w=800&h=400&fit=crop',
      color: 'blue'
    },
    {
      title: 'Dental Clinic - 5x Patient Growth',
      description: 'Increased monthly patient appointments from 50 to 250+ through local SEO and Google Ads',
      metrics: '5x Growth',
      image: 'https://images.unsplash.com/photo-1629909613654-28e377c37b09?w=800&h=400&fit=crop',
      color: 'green'
    },
    {
      title: 'E-commerce Store - 300% Sales Boost',
      description: 'Tripled monthly sales for fashion e-commerce store through strategic Facebook & Instagram ads',
      metrics: '300% Increase',
      image: 'https://images.unsplash.com/photo-1472851294608-062f824d29cc?w=800&h=400&fit=crop',
      color: 'purple'
    }
  ];

  return (
    <>
      <Helmet>
        <title>Best IT Training Institute & Digital Marketing Agency in Malad West, Mumbai - CM Techno Solution</title>
        <meta 
          name="description" 
          content="CM Techno Solution - Leading IT training institute and digital marketing agency in Malad West, Mumbai. Offering Python, Full Stack, Digital Marketing courses and performance marketing services including real estate lead generation and local business marketing." 
        />
      </Helmet>

      {/* Hero Section */}
      <section className="relative min-h-screen flex items-center justify-center overflow-hidden">
        {/* Background Image with Overlay */}
        <div className="absolute inset-0 z-0">
          <img
            src="https://images.unsplash.com/photo-1603985585179-3d71c35a537c"
            alt="Modern tech workspace with team collaboration"
            className="w-full h-full object-cover"
          />
          <div className="absolute inset-0 bg-gradient-to-br from-blue-900/95 via-blue-800/90 to-blue-900/95"></div>
        </div>

        {/* Content */}
        <div className="relative z-10 container mx-auto px-4 py-20 text-center">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
          >
            <div className="inline-block px-4 py-2 bg-white/10 backdrop-blur-sm rounded-full border border-white/20 mb-6">
              <span className="text-white text-sm font-medium">🚀 Empowering Careers & Growing Businesses</span>
            </div>
            
            <h1 className="text-4xl md:text-5xl lg:text-6xl font-bold text-white mb-6 leading-tight">
              Best <span className="text-red-400">IT</span> Training Institute & <br />
              Digital Marketing Agency in <br />
              <span className="text-blue-300">Malad West, Mumbai</span>
            </h1>
            
            <p className="text-xl md:text-2xl text-blue-100 mb-8 max-w-3xl mx-auto">
              Professional Courses, Real Estate Lead Generation & Performance Marketing Services
            </p>

            <div className="flex flex-col sm:flex-row gap-4 justify-center items-center">
              <Button
                onClick={onOpenEnquiry}
                className="px-8 py-6 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-2xl"
              >
                Enquire for Course
              </Button>
              <Link to="/digital-marketing-services">
                <Button className="px-8 py-6 bg-green-600 hover:bg-green-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-2xl">
                  Get Business Leads
                </Button>
              </Link>
              <Link to="/contact">
                <Button className="px-8 py-6 bg-white hover:bg-gray-100 text-blue-900 rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-2xl">
                  Book Free Consultation
                </Button>
              </Link>
            </div>
          </motion.div>
        </div>

        {/* Scroll Indicator */}
        <div className="absolute bottom-8 left-1/2 transform -translate-x-1/2 animate-bounce">
          <div className="w-6 h-10 border-2 border-white/50 rounded-full flex justify-center">
            <div className="w-1 h-3 bg-white rounded-full mt-2"></div>
          </div>
        </div>
      </section>

      {/* Stats Section */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
            {stats.map((stat, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="text-center p-6 bg-gradient-to-br from-blue-50 to-blue-100 rounded-2xl shadow-lg hover:shadow-xl transition-shadow"
              >
                <stat.icon className="w-12 h-12 mx-auto mb-3 text-blue-600" />
                <div className="text-3xl font-bold text-blue-900 mb-2">{stat.value}</div>
                <div className="text-sm text-gray-700 font-medium">{stat.label}</div>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Services Highlights */}
      <section className="py-16 bg-gradient-to-br from-gray-50 to-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Our Services
            </h2>
            <p className="text-xl text-gray-600 max-w-2xl mx-auto">
              Comprehensive IT training and digital marketing solutions to help you succeed
            </p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {services.map((service, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
              >
                <Link to={service.link}>
                  <div className="p-6 bg-white rounded-2xl shadow-lg hover:shadow-2xl transition-all hover:scale-105 h-full border border-gray-100">
                    <div className={`w-14 h-14 bg-${service.color}-100 rounded-xl flex items-center justify-center mb-4`}>
                      <service.icon className={`w-7 h-7 text-${service.color}-600`} />
                    </div>
                    <h3 className="text-xl font-bold text-gray-900 mb-3">{service.title}</h3>
                    <p className="text-gray-600 mb-4">{service.description}</p>
                    <div className="flex items-center text-blue-600 font-semibold">
                      Learn More <ArrowRight className="w-4 h-4 ml-2" />
                    </div>
                  </div>
                </Link>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Testimonials Section */}
      <section className="py-16 bg-gradient-to-br from-blue-900 to-blue-800 text-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold mb-4">What Our Clients Say</h2>
            <p className="text-xl text-blue-100">Real success stories from our students and clients</p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            {testimonials.map((testimonial, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-6 bg-white/10 backdrop-blur-sm rounded-2xl border border-white/20 hover:bg-white/20 transition-all"
              >
                <div className="flex items-center mb-4">
                  <img
                    src={testimonial.image}
                    alt={testimonial.name}
                    className="w-12 h-12 rounded-full object-cover mr-3"
                  />
                  <div>
                    <div className="font-bold">{testimonial.name}</div>
                    <div className="text-sm text-blue-200">{testimonial.role}</div>
                  </div>
                </div>
                <div className="flex mb-3">
                  {[...Array(testimonial.rating)].map((_, i) => (
                    <Star key={i} className="w-4 h-4 fill-yellow-400 text-yellow-400" />
                  ))}
                </div>
                <p className="text-blue-100 text-sm">{testimonial.text}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Case Studies Section */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">Success Stories</h2>
            <p className="text-xl text-gray-600">Real results from our clients</p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {caseStudies.map((study, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="rounded-2xl overflow-hidden shadow-lg hover:shadow-2xl transition-all hover:scale-105"
              >
                <img
                  src={study.image}
                  alt={study.title}
                  className="w-full h-48 object-cover"
                />
                <div className="p-6 bg-white">
                  <div className={`inline-block px-3 py-1 bg-${study.color}-100 text-${study.color}-700 rounded-full text-sm font-semibold mb-3`}>
                    {study.metrics}
                  </div>
                  <h3 className="text-xl font-bold text-gray-900 mb-2">{study.title}</h3>
                  <p className="text-gray-600">{study.description}</p>
                </div>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-gradient-to-r from-blue-600 to-blue-800 text-white">
        <div className="container mx-auto px-4 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
          >
            <h2 className="text-4xl font-bold mb-4">
              Ready to Transform Your Career or Business?
            </h2>
            <p className="text-xl text-blue-100 mb-8 max-w-2xl mx-auto">
              Join 500+ successful students and 100+ satisfied business clients. Start your journey today!
            </p>
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button
                onClick={onOpenEnquiry}
                className="px-8 py-6 bg-white hover:bg-gray-100 text-blue-900 rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl"
              >
                Get Started Now
              </Button>
              <Link to="/contact">
                <Button className="px-8 py-6 bg-blue-500 hover:bg-blue-600 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl border-2 border-white">
                  Talk to Expert
                </Button>
              </Link>
            </div>
          </motion.div>
        </div>
      </section>
    </>
  );
}

export default HomePage;
