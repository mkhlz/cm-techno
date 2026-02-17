import React from 'react';
import { Link } from 'react-router-dom';
import { MapPin, Mail, Phone, MessageCircle, Facebook, Instagram, Linkedin, Twitter } from 'lucide-react';

function Footer() {
  const quickLinks = [
    { name: 'Home', path: '/' },
    { name: 'IT Courses', path: '/it-courses' },
    { name: 'Franchise', path: '/franchise' },
    { name: 'Digital Services', path: '/digital-marketing-services' },
    { name: 'Why Choose Us', path: '/why-choose-us' },
    { name: 'Contact', path: '/contact' }
  ];

  return (
    <footer className="bg-white mx-10 text-gray-800">
      <div className="container mx-auto px-4 py-12">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">

          {/* Company Info */}
          <div>
            <div className="flex items-center mb-4">
              <img
                src="/assets/cm-techno-logo.png"   // <-- transparent logo
                alt="CM Techno Solution Logo"
                className="h-16 w-auto object-contain"
              />
            </div>

            <p className="text-gray-600 mb-4 text-sm">
              Leading IT Training Institute and Digital Marketing Agency in
              Mumbai, providing quality education and ROI-focused marketing
              solutions.
            </p>

            <div className="space-y-2 text-sm">
              <div className="flex items-start space-x-2">
                <MapPin className="w-4 h-4 mt-1 flex-shrink-0 text-blue-600" />
                <span className="text-gray-600">
                  A-17, 1st Floor Patel Shopping Center, Near Malad Subway, Opp.
                  Foodland Hotel, Malad (West), Mumbai-400064
                </span>
              </div>

              <div className="flex items-center space-x-2">
                <Mail className="w-4 h-4 flex-shrink-0 text-blue-600" />
                <a
                  href="mailto:cmskillindia@gmail.com"
                  className="text-gray-600 hover:text-blue-600 transition-colors"
                >
                  cmskillindia@gmail.com
                </a>
              </div>

              <div className="flex items-center space-x-2">
                <Phone className="w-4 h-4 flex-shrink-0 text-blue-600" />
                <a
                  href="tel:+918169809775"
                  className="text-gray-600 hover:text-blue-600 transition-colors"
                >
                  +91 81698 09775
                </a>
              </div>

              <div className="flex items-center space-x-2">
                <MessageCircle className="w-4 h-4 flex-shrink-0 text-blue-600" />
                <a
                  href="https://wa.me/918169809775?text=Hi%20CM%20Techno%20Solution"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-gray-600 hover:text-green-600 transition-colors"
                >
                  WhatsApp Chat
                </a>
              </div>
            </div>
          </div>

          {/* Quick Links */}
          <div>
            <h3 className="text-lg font-bold mb-4 text-gray-900">Quick Links</h3>
            <ul className="space-y-2">
              {quickLinks.map((link) => (
                <li key={link.path}>
                  <Link
                    to={link.path}
                    className="text-gray-600 hover:text-blue-600 transition-all text-sm block hover:translate-x-1"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          {/* Our Services */}
          <div>
            <h3 className="text-lg font-bold mb-4 text-gray-900">Our Services</h3>
            <ul className="space-y-2 text-sm text-gray-600">
              <li>IT Training Courses</li>
              <li>Real Estate Lead Generation</li>
              <li>Local Business Marketing</li>
              <li>E-commerce Marketing</li>
              <li>Website Development</li>
              <li>SEO Services</li>
              <li>Social Media Marketing</li>
              <li>Google & Meta Ads</li>
            </ul>
          </div>

          {/* Connect With Us */}
          <div>
            <h3 className="text-lg font-bold mb-4 text-gray-900">
              Connect With Us
            </h3>

            <p className="text-gray-600 mb-4 text-sm">
              Follow us on social media for updates, tips, and industry
              insights.
            </p>

            <div className="flex space-x-3 mb-6">
              <a
                href="#"
                className="w-10 h-10 bg-gray-200 hover:bg-blue-600 hover:text-white rounded-full flex items-center justify-center transition-all hover:scale-110"
                aria-label="Facebook"
              >
                <Facebook className="w-5 h-5" />
              </a>

              <a
                href="#"
                className="w-10 h-10 bg-gray-200 hover:bg-pink-600 hover:text-white rounded-full flex items-center justify-center transition-all hover:scale-110"
                aria-label="Instagram"
              >
                <Instagram className="w-5 h-5" />
              </a>

              <a
                href="#"
                className="w-10 h-10 bg-gray-200 hover:bg-blue-700 hover:text-white rounded-full flex items-center justify-center transition-all hover:scale-110"
                aria-label="LinkedIn"
              >
                <Linkedin className="w-5 h-5" />
              </a>

              <a
                href="#"
                className="w-10 h-10 bg-gray-200 hover:bg-sky-500 hover:text-white rounded-full flex items-center justify-center transition-all hover:scale-110"
                aria-label="Twitter"
              >
                <Twitter className="w-5 h-5" />
              </a>
            </div>

            <Link to="/franchise">
              <button className="px-4 py-2 bg-gradient-to-r from-red-600 to-red-500 hover:from-red-500 hover:to-red-600 rounded-lg text-sm font-bold text-white shadow-lg transition-all hover:scale-105 w-full">
                Become a Franchise Partner
              </button>
            </Link>
          </div>
        </div>

        {/* Bottom Bar */}
        <div className="border-t border-gray-200 mt-8 pt-8 text-center">
          <p className="text-gray-600 text-sm">
            © {new Date().getFullYear()} CM Techno Solution. All rights reserved.
          </p>
          <p className="text-gray-500 text-xs mt-2">
            Empowering Careers & Growing Businesses Since 2002
          </p>
        </div>
      </div>
    </footer>
  );
}

export default Footer;