
import React, { useState } from 'react';
import { Helmet } from 'react-helmet';
import { motion } from 'framer-motion';
import { 
  MapPin, Mail, Phone, MessageCircle, 
  Clock, Send, Building2, User
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useToast } from '@/components/ui/use-toast';

function ContactPage() {
  const { toast } = useToast();
  const [generalFormData, setGeneralFormData] = useState({
    name: '',
    phone: '',
    email: '',
    service: '',
    message: ''
  });

  const [ businessFormData, setBusinessFormData] = useState({
    companyName: '',
    contactPerson: '',
    phone: '',
    email: '',
    businessType: '',
    budget: '',
    message: ''
  });

  const [isSubmittingGeneral, setIsSubmittingGeneral] = useState(false);
  const [isSubmittingBusiness, setIsSubmittingBusiness] = useState(false);

  const handleGeneralSubmit = async (e) => {
    e.preventDefault();
    setIsSubmittingGeneral(true);

    try {
      const response = await fetch('https://formspree.io/f/xanydpvb', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...generalFormData,
          _subject: `General Enquiry from ${generalFormData.name}`
        })
      });

      if (response.ok) {
        toast({
          title: "Success!",
          description: "Your enquiry has been submitted. We'll contact you soon!",
        });
        setGeneralFormData({ name: '', phone: '', email: '', service: '', message: '' });
      }
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to submit form. Please try again.",
        variant: "destructive"
      });
    } finally {
      setIsSubmittingGeneral(false);
    }
  };

  const handleBusinessSubmit = async (e) => {
    e.preventDefault();
    setIsSubmittingBusiness(true);

    try {
      const response = await fetch('https://formspree.io/f/xanydpvb', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...businessFormData,
          _subject: `Business Consultation Request from ${businessFormData.companyName}`
        })
      });

      if (response.ok) {
        toast({
          title: "Success!",
          description: "Your consultation request has been submitted. We'll contact you soon!",
        });
        setBusinessFormData({ companyName: '', contactPerson: '', phone: '', email: '', businessType: '', budget: '', message: '' });
      }
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to submit form. Please try again.",
        variant: "destructive"
      });
    } finally {
      setIsSubmittingBusiness(false);
    }
  };

  return (
    <>
      <Helmet>
        <title>Contact CM Techno Solution - IT Training & Digital Marketing Agency in Malad West, Mumbai</title>
        <meta 
          name="description" 
          content="Contact CM Techno Solution for IT courses and digital marketing services. Located in Malad West, Mumbai. Call +91 81698 09775 or email cmskillindia@gmail.com" 
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
            <h1 className="text-4xl md:text-5xl lg:text-6xl font-bold mb-6">
              Contact <span className="text-blue-300">CM Techno Solution</span>
            </h1>
            
            <p className="text-xl md:text-2xl text-blue-100 mb-8">
              Get in touch with us for IT training courses or digital marketing services
            </p>
          </motion.div>
        </div>
      </section>

      {/* Contact Information */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-12">
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              className="p-6 bg-gradient-to-br from-blue-50 to-blue-100 rounded-2xl shadow-lg text-center"
            >
              <MapPin className="w-12 h-12 mx-auto mb-4 text-blue-600" />
              <h3 className="font-bold text-gray-900 mb-2">Address</h3>
              <p className="text-gray-700 text-sm">
                A-17, 1st Floor Patel Shopping Center, Near Malad Subway, Opp. Foodland Hotel, Malad (West), Mumbai-400064
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              transition={{ delay: 0.1 }}
              viewport={{ once: true }}
              className="p-6 bg-gradient-to-br from-green-50 to-green-100 rounded-2xl shadow-lg text-center"
            >
              <Mail className="w-12 h-12 mx-auto mb-4 text-green-600" />
              <h3 className="font-bold text-gray-900 mb-2">Email</h3>
              <a href="mailto:cmskillindia@gmail.com" className="text-gray-700 text-sm hover:text-green-600">
                cmskillindia@gmail.com
              </a>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              transition={{ delay: 0.2 }}
              viewport={{ once: true }}
              className="p-6 bg-gradient-to-br from-purple-50 to-purple-100 rounded-2xl shadow-lg text-center"
            >
              <Phone className="w-12 h-12 mx-auto mb-4 text-purple-600" />
              <h3 className="font-bold text-gray-900 mb-2">Phone</h3>
              <a href="tel:+918169809775" className="text-gray-700 text-sm hover:text-purple-600">
                +91 81698 09775
              </a>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              transition={{ delay: 0.3 }}
              viewport={{ once: true }}
              className="p-6 bg-gradient-to-br from-orange-50 to-orange-100 rounded-2xl shadow-lg text-center"
            >
              <Clock className="w-12 h-12 mx-auto mb-4 text-orange-600" />
              <h3 className="font-bold text-gray-900 mb-2">Hours</h3>
              <p className="text-gray-700 text-sm">
                Mon - Sat: 9 AM - 8 PM<br />
                Sunday: Closed
              </p>
            </motion.div>
          </div>

          {/* Quick Contact Buttons */}
          <div className="flex flex-col sm:flex-row gap-4 justify-center mb-12">
            <a href="tel:+918169809775">
              <Button className="px-8 py-6 bg-green-600 hover:bg-green-700 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl">
                <Phone className="w-5 h-5 mr-2" />
                Call Now
              </Button>
            </a>
            <a 
              href="https://wa.me/918169809775?text=Hi%20CM%20Techno%20Solution"
              target="_blank"
              rel="noopener noreferrer"
            >
              <Button className="px-8 py-6 bg-green-500 hover:bg-green-600 text-white rounded-xl font-semibold text-lg transition-all hover:scale-105 shadow-xl">
                <MessageCircle className="w-5 h-5 mr-2" />
                WhatsApp Chat
              </Button>
            </a>
          </div>
        </div>
      </section>

      {/* Google Map */}
      <section className="py-16 bg-gray-50">
        <div className="container mx-auto px-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            className="rounded-2xl overflow-hidden shadow-2xl"
          >
            <iframe
              src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3768.4417182285!2d72.8465063!3d19.1843636!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x3be7b6f6f6f6f6f7%3A0x6f6f6f6f6f6f6f6f!2sPatel%20Shopping%20Center%2C%20Malad%20West%2C%20Mumbai%2C%20Maharashtra%20400064!5e0!3m2!1sen!2sin!4v1645000000000!5m2!1sen!2sin"
              width="100%"
              height="450"
              style={{ border: 0 }}
              allowFullScreen=""
              loading="lazy"
              title="CM Techno Solution Location - Malad West, Mumbai"
            ></iframe>
          </motion.div>
        </div>
      </section>

      {/* Contact Forms */}
      <section className="py-16 bg-white">
        <div className="container mx-auto px-4">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {/* General Enquiry Form */}
            <motion.div
              initial={{ opacity: 0, x: -20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
              className="bg-gradient-to-br from-blue-50 to-white p-8 rounded-2xl shadow-xl"
            >
              <div className="flex items-center mb-6">
                <User className="w-8 h-8 text-blue-600 mr-3" />
                <h2 className="text-3xl font-bold text-gray-900">General Enquiry</h2>
              </div>
              
              <form onSubmit={handleGeneralSubmit} className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Full Name *
                  </label>
                  <input
                    type="text"
                    value={generalFormData.name}
                    onChange={(e) => setGeneralFormData({ ...generalFormData, name: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    placeholder="Enter your name"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Phone Number *
                  </label>
                  <input
                    type="tel"
                    value={generalFormData.phone}
                    onChange={(e) => setGeneralFormData({ ...generalFormData, phone: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    placeholder="10-digit mobile number"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Email Address *
                  </label>
                  <input
                    type="email"
                    value={generalFormData.email}
                    onChange={(e) => setGeneralFormData({ ...generalFormData, email: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    placeholder="your.email@example.com"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Service Required *
                  </label>
                  <select
                    value={generalFormData.service}
                    onChange={(e) => setGeneralFormData({ ...generalFormData, service: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    required
                  >
                    <option value="">Select a service</option>
                    <option value="IT Course">IT Course</option>
                    <option value="Real Estate Leads">Real Estate Leads</option>
                    <option value="Local Business Marketing">Local Business Marketing</option>
                    <option value="E-commerce Marketing">E-commerce Marketing</option>
                    <option value="Website Development">Website Development</option>
                    <option value="Other">Other</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Message
                  </label>
                  <textarea
                    value={generalFormData.message}
                    onChange={(e) => setGeneralFormData({ ...generalFormData, message: e.target.value })}
                    rows="4"
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    placeholder="Tell us more about your requirements..."
                  ></textarea>
                </div>

                <Button
                  type="submit"
                  disabled={isSubmittingGeneral}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-lg font-semibold transition-all hover:scale-105"
                >
                  {isSubmittingGeneral ? 'Submitting...' : (
                    <>
                      <Send className="w-5 h-5 mr-2 inline" />
                      Submit Enquiry
                    </>
                  )}
                </Button>
              </form>
            </motion.div>

            {/* Business Consultation Form */}
            <motion.div
              initial={{ opacity: 0, x: 20 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
              className="bg-gradient-to-br from-green-50 to-white p-8 rounded-2xl shadow-xl"
            >
              <div className="flex items-center mb-6">
                <Building2 className="w-8 h-8 text-green-600 mr-3" />
                <h2 className="text-3xl font-bold text-gray-900">Business Consultation</h2>
              </div>
              
              <form onSubmit={handleBusinessSubmit} className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Company Name *
                  </label>
                  <input
                    type="text"
                    value={businessFormData.companyName}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, companyName: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    placeholder="Your company name"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Contact Person *
                  </label>
                  <input
                    type="text"
                    value={businessFormData.contactPerson}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, contactPerson: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    placeholder="Your name"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Phone Number *
                  </label>
                  <input
                    type="tel"
                    value={businessFormData.phone}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, phone: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    placeholder="10-digit mobile number"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Email Address *
                  </label>
                  <input
                    type="email"
                    value={businessFormData.email}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, email: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    placeholder="business@company.com"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Business Type *
                  </label>
                  <select
                    value={businessFormData.businessType}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, businessType: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    required
                  >
                    <option value="">Select business type</option>
                    <option value="Real Estate">Real Estate</option>
                    <option value="Healthcare">Healthcare</option>
                    <option value="Fitness">Fitness</option>
                    <option value="Beauty & Wellness">Beauty & Wellness</option>
                    <option value="E-commerce">E-commerce</option>
                    <option value="Professional Services">Professional Services</option>
                    <option value="Other">Other</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Monthly Budget
                  </label>
                  <select
                    value={businessFormData.budget}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, budget: e.target.value })}
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                  >
                    <option value="">Select budget range</option>
                    <option value="Under ₹25,000">Under ₹25,000</option>
                    <option value="₹25,000 - ₹50,000">₹25,000 - ₹50,000</option>
                    <option value="₹50,000 - ₹1,00,000">₹50,000 - ₹1,00,000</option>
                    <option value="Above ₹1,00,000">Above ₹1,00,000</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Message
                  </label>
                  <textarea
                    value={businessFormData.message}
                    onChange={(e) => setBusinessFormData({ ...businessFormData, message: e.target.value })}
                    rows="4"
                    className="w-full px-4 py-3 bg-white border border-gray-300 rounded-lg text-gray-900 placeholder:text-gray-500 focus:ring-2 focus:ring-green-500 focus:border-transparent"
                    placeholder="Tell us about your business goals..."
                  ></textarea>
                </div>

                <Button
                  type="submit"
                  disabled={isSubmittingBusiness}
                  className="w-full bg-green-600 hover:bg-green-700 text-white py-3 rounded-lg font-semibold transition-all hover:scale-105"
                >
                  {isSubmittingBusiness ? 'Submitting...' : (
                    <>
                      <Send className="w-5 h-5 mr-2 inline" />
                      Request Consultation
                    </>
                  )}
                </Button>
              </form>
            </motion.div>
          </div>
        </div>
      </section>
    </>
  );
}

export default ContactPage;
