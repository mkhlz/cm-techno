
import React from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { 
  Building2, Stethoscope, Dumbbell, Scissors, Calculator,
  ShoppingBag, TrendingUp, Search, MessageCircle, Globe,
  BarChart3, Target, Zap, CheckCircle, ArrowRight
} from 'lucide-react';
import { Button } from '@/components/ui/button';

function DigitalMarketingServicesPage({ onOpenEnquiry }) {
  const realEstateServices = [
    { title: 'Ready-to-Move Properties', desc: 'Targeted campaigns for immediate buyers' },
    { title: 'Under Construction Projects', desc: 'Long-term investment focused leads' },
    { title: 'Channel Partner Marketing', desc: 'Builder collaboration campaigns' },
    { title: 'High-Intent Buyer Leads', desc: 'Quality leads ready to purchase' }
  ];

  const localBusinessServices = [
    {
      icon: Stethoscope,
      title: 'Doctors & Clinics',
      description: 'Patient acquisition campaigns with geo-targeted ads',
      color: 'red'
    },
    {
      icon: Dumbbell,
      title: 'Gyms & Fitness Centers',
      description: 'Membership drive campaigns with local targeting',
      color: 'green'
    },
    {
      icon: Scissors,
      title: 'Salons & Beauty Clinics',
      description: 'Appointment booking campaigns with offers',
      color: 'purple'
    },
    {
      icon: Calculator,
      title: 'CA & Professional Services',
      description: 'Lead generation for consultancy services',
      color: 'blue'
    }
  ];

  const digitalServices = [
    {
      icon: ShoppingBag,
      title: 'E-commerce Marketing',
      description: 'Shopify marketing, Facebook & Instagram ads, conversion optimization',
      features: ['Product Catalog Ads', 'Retargeting Campaigns', 'Conversion Tracking']
    },
    {
      icon: TrendingUp,
      title: 'Social Media Marketing',
      description: 'Complete social media management with content creation and engagement',
      features: ['Content Strategy', 'Daily Posting', 'Community Management']
    },
    {
      icon: Target,
      title: 'Google Ads & Meta Ads',
      description: 'Performance marketing campaigns with measurable ROI',
      features: ['Search Ads', 'Display Ads', 'Shopping Ads']
    },
    {
      icon: Globe,
      title: 'Website Development',
      description: 'Professional websites optimized for conversions and SEO',
      features: ['Responsive Design', 'SEO Optimized', 'Fast Loading']
    },
    {
      icon: Search,
      title: 'SEO Services',
      description: 'Rank higher on Google with our proven SEO strategies',
      features: ['On-Page SEO', 'Off-Page SEO', 'Technical SEO']
    },
    {
      icon: MessageCircle,
      title: 'WhatsApp Marketing',
      description: 'Direct customer engagement through WhatsApp campaigns',
      features: ['Broadcast Messages', 'Automated Replies', 'Customer Support']
    }
  ];

  const results = [
    { metric: '₹50 Cr+', label: 'Real Estate Sales Generated' },
    { metric: '10,000+', label: 'Quality Leads Delivered' },
    { metric: '5x', label: 'Average ROI for Clients' },
    { metric: '100+', label: 'Successful Campaigns' }
  ];

  return (
    <>
      <Helmet>
        <title>Digital Marketing Agency in Mumbai - Lead Generation & Performance Marketing | CM Techno Solution</title>
        <meta 
          name="description" 
          content="Top digital marketing agency in Mumbai specializing in real estate lead generation, local business marketing, e-commerce marketing, Google Ads, Facebook Ads, SEO, and website development. ROI-focused performance marketing services." 
        />
      </Helmet>

      {/* Hero Section */}
      <section className="relative pt-40 pb-16 bg-gradient-to-br from-blue-900 via-blue-800 to-blue-900 text-white overflow-hidden">
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
              <span className="text-sm font-medium">🚀 ROI-Focused Performance Marketing</span>
            </div>
            
            <h1 className="text-4xl md:text-5xl lg:text-6xl font-bold mb-6 leading-tight">
              Digital Marketing Agency in Mumbai
            </h1>
            
            <p className="text-xl md:text-2xl text-blue-100 mb-8">
              Lead Generation & Performance Marketing for Real Estate, Local Businesses & E-commerce
            </p>

            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button
                onClick={onOpenEnquiry}
                className="px-8 py-6 bg-white hover:bg-gray-100 text-blue-900 rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-2xl"
              >
                Get Started
              </Button>
              <a href="tel:+919876543210">
                <Button className="px-8 py-6 bg-green-600 hover:bg-green-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-2xl">
                  Call Now
                </Button>
              </a>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Results Section */}
      <section className="py-12 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
            {results.map((result, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="text-center p-6 bg-gradient-to-br from-blue-50 to-blue-100 rounded-2xl"
              >
                <div className="text-3xl font-bold text-blue-600 mb-2">{result.metric}</div>
                <div className="text-sm text-gray-700 font-medium">{result.label}</div>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Real Estate Lead Generation */}
      <section className="py-16 bg-gradient-to-br from-gray-50 to-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Real Estate Lead Generation
            </h2>
            <p className="text-xl text-gray-600 max-w-3xl mx-auto">
              High-intent buyer leads for residential and commercial properties
            </p>
          </motion.div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-center mb-12">
            <motion.div
              initial={{ opacity: 0, x: -20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <img
                src="https://images.unsplash.com/photo-1698316738298-7f92b28225e4?w=800&h=600&fit=crop"
                alt="Real estate lead generation campaigns showing property marketing"
                className="rounded-2xl shadow-2xl"
              />
            </motion.div>

            <motion.div
              initial={{ opacity: 0, x: 20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <div className="space-y-4">
                {realEstateServices.map((service, index) => (
                  <div key={index} className="p-4 bg-white rounded-xl shadow-md hover:shadow-lg transition-shadow">
                    <div className="flex items-start">
                      <CheckCircle className="w-6 h-6 text-green-600 mr-3 flex-shrink-0 mt-1" />
                      <div>
                        <h3 className="font-bold text-gray-900 mb-1">{service.title}</h3>
                        <p className="text-gray-600 text-sm">{service.desc}</p>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              <div className="mt-6 p-6 bg-blue-50 rounded-xl border border-blue-200">
                <div className="flex items-start">
                  <Building2 className="w-8 h-8 text-blue-600 mr-4 flex-shrink-0" />
                  <div>
                    <h4 className="font-bold text-gray-900 mb-2">₹50 Cr+ Sales Generated</h4>
                    <p className="text-gray-700 text-sm">
                      We've helped real estate developers and channel partners generate over ₹50 Crore in property sales through targeted digital marketing campaigns.
                    </p>
                  </div>
                </div>
              </div>

              <Button
                onClick={onOpenEnquiry}
                className="w-full mt-6 bg-blue-600 hover:bg-blue-700 text-white py-4 rounded-xl font-semibold text-lg transition-all hover:scale-105"
              >
                Get Real Estate Leads
              </Button>
            </motion.div>
          </div>
        </div>
      </section>

      {/* Local Business Marketing */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Local Business Lead Generation
            </h2>
            <p className="text-xl text-gray-600 max-w-3xl mx-auto">
              Targeted campaigns for service-based businesses in Mumbai
            </p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            {localBusinessServices.map((service, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-6 bg-gradient-to-br from-white to-gray-50 rounded-2xl shadow-lg hover:shadow-2xl transition-all hover:scale-105 border border-gray-200"
              >
                <div className={`w-16 h-16 bg-${service.color}-100 rounded-2xl flex items-center justify-center mb-4`}>
                  <service.icon className={`w-8 h-8 text-${service.color}-600`} />
                </div>
                <h3 className="text-xl font-bold text-gray-900 mb-2">{service.title}</h3>
                <p className="text-gray-600">{service.description}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Digital Marketing Services */}
      <section className="py-16 bg-gradient-to-br from-gray-50 to-white">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="text-center mb-12"
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Complete Digital Marketing Solutions
            </h2>
            <p className="text-xl text-gray-600 max-w-3xl mx-auto">
              End-to-end digital marketing services to grow your business
            </p>
          </motion.div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
            {digitalServices.map((service, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                viewport={{ once: true }}
                className="p-6 bg-white rounded-2xl shadow-lg hover:shadow-2xl transition-all hover:scale-105"
              >
                <div className="w-14 h-14 bg-blue-100 rounded-xl flex items-center justify-center mb-4">
                  <service.icon className="w-7 h-7 text-blue-600" />
                </div>
                <h3 className="text-xl font-bold text-gray-900 mb-2">{service.title}</h3>
                <p className="text-gray-600 mb-4">{service.description}</p>
                <ul className="space-y-2">
                  {service.features.map((feature, i) => (
                    <li key={i} className="flex items-center text-sm text-gray-700">
                      <CheckCircle className="w-4 h-4 text-green-600 mr-2 flex-shrink-0" />
                      {feature}
                    </li>
                  ))}
                </ul>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Service Image Section */}
      <section className="py-16 bg-gradient-to-r from-blue-600 to-blue-800 text-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center">
            <motion.div
              initial={{ opacity: 0, x: -20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <h2 className="text-4xl font-bold mb-6">
                Performance Marketing with Proven Results
              </h2>
              <p className="text-xl text-blue-100 mb-6">
                We focus on ROI, not vanity metrics. Every campaign is designed to generate measurable business results.
              </p>
              <ul className="space-y-4">
                <li className="flex items-center">
                  <Zap className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Data-driven campaign strategies</span>
                </li>
                <li className="flex items-center">
                  <BarChart3 className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Transparent reporting and analytics</span>
                </li>
                <li className="flex items-center">
                  <Target className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Continuous optimization for better ROI</span>
                </li>
                <li className="flex items-center">
                  <CheckCircle className="w-6 h-6 mr-3 flex-shrink-0" />
                  <span>Dedicated account manager</span>
                </li>
              </ul>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, x: 20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
            >
              <img
                src="https://images.unsplash.com/photo-1603985585179-3d71c35a537c?w=800&h=600&fit=crop"
                alt="Digital marketing agency team working on performance marketing campaigns"
                className="rounded-2xl shadow-2xl"
              />
            </motion.div>
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
          >
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Ready to Grow Your Business?
            </h2>
            <p className="text-xl text-gray-600 mb-8 max-w-2xl mx-auto">
              Get a free marketing consultation and discover how we can help you achieve your business goals
            </p>
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button
                onClick={onOpenEnquiry}
                className="px-8 py-6 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl"
              >
                Get Free Consultation
              </Button>
              <a href="https://wa.me/919876543210?text=Hi%2C%20I%20want%20to%20discuss%20digital%20marketing%20services">
                <Button className="px-8 py-6 bg-green-600 hover:bg-green-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl">
                  WhatsApp Us
                </Button>
              </a>
            </div>
          </motion.div>
        </div>
      </section>
    </>
  );
}

export default DigitalMarketingServicesPage;
