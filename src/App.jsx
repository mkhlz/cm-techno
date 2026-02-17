
import React, { useState, useEffect } from 'react';
import { Route, Routes, BrowserRouter as Router, useLocation } from 'react-router-dom';
import { MessageCircle } from 'lucide-react';
import ScrollToTop from '@/components/ScrollToTop';
import Header from '@/components/Header';
import Footer from '@/components/Footer';
import EnquiryFormModal from '@/components/EnquiryFormModal';
import HomePage from '@/pages/HomePage';
import ITCoursesPage from '@/pages/ITCoursesPage';
import FranchisePage from '@/pages/FranchisePage';
import CourseDetailsTemplate from '@/pages/CourseDetailsTemplate';
import DigitalMarketingServicesPage from '@/pages/DigitalMarketingServicesPage';
import WhyChooseUsPage from '@/pages/WhyChooseUsPage';
import ContactPage from '@/pages/ContactPage';
import { Toaster } from '@/components/ui/toaster';

function AppContent() {
  const [isFormModalOpen, setIsFormModalOpen] = useState(false);
  const location = useLocation();

  const handleOpenEnquiry = () => {
    setIsFormModalOpen(true);
  };

  const handleCloseEnquiry = () => {
    setIsFormModalOpen(false);
  };

  // Auto-open enquiry modal 5 seconds after page load
  useEffect(() => {
    const timer = setTimeout(() => {
      setIsFormModalOpen(true);
    }, 5000);
    return () => clearTimeout(timer);
  }, [location.pathname]);

  return (
    <>
      <ScrollToTop />
      <div className="min-h-screen flex flex-col bg-white">
        <Header onOpenEnquiry={handleOpenEnquiry} />
        <main className="flex-grow">
          <Routes>
            <Route path="/" element={<HomePage onOpenEnquiry={handleOpenEnquiry} />} />
            <Route path="/it-courses" element={<ITCoursesPage onOpenEnquiry={handleOpenEnquiry} />} />
            <Route path="/courses/:slug" element={<CourseDetailsTemplate onOpenEnquiry={handleOpenEnquiry} />} />
            <Route path="/franchise" element={<FranchisePage />} />
            <Route path="/digital-marketing-services" element={<DigitalMarketingServicesPage onOpenEnquiry={handleOpenEnquiry} />} />
            <Route path="/why-choose-us" element={<WhyChooseUsPage onOpenEnquiry={handleOpenEnquiry} />} />
            <Route path="/contact" element={<ContactPage />} />
          </Routes>
        </main>
        <Footer />
        <EnquiryFormModal isOpen={isFormModalOpen} onClose={handleCloseEnquiry} />
        <Toaster />
      </div>

      {/* Floating WhatsApp Chat Button */}
      <a
        href="https://wa.me/918169809775?text=Hi%20CM%20Techno%20Solution"
        target="_blank"
        rel="noopener noreferrer"
        className="fixed bottom-4 right-4 z-50 bg-green-500 hover:bg-green-600 text-white p-4 rounded-full shadow-lg transition-transform hover:scale-110"
        aria-label="Chat on WhatsApp"
      >
        <MessageCircle className="w-6 h-6" />
      </a>
    </>
  );
}

function App() {
  return (
    <Router>
      <AppContent />
    </Router>
  );
}

export default App;
