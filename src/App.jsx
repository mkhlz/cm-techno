
import React, { useState } from 'react';
import { Route, Routes, BrowserRouter as Router } from 'react-router-dom';
import ScrollToTop from '@/components/ScrollToTop';
import Header from '@/components/Header';
import Footer from '@/components/Footer';
import FormModal from '@/components/FormModal';
import HomePage from '@/pages/HomePage';
import ITCoursesPage from '@/pages/ITCoursesPage';
import FranchisePage from '@/pages/FranchisePage';
import CourseDetailsTemplate from '@/pages/CourseDetailsTemplate';
import DigitalMarketingServicesPage from '@/pages/DigitalMarketingServicesPage';
import WhyChooseUsPage from '@/pages/WhyChooseUsPage';
import ContactPage from '@/pages/ContactPage';
import { Toaster } from '@/components/ui/toaster';

function App() {
  const [isFormModalOpen, setIsFormModalOpen] = useState(false);

  const handleOpenEnquiry = () => {
    setIsFormModalOpen(true);
  };

  const handleCloseEnquiry = () => {
    setIsFormModalOpen(false);
  };

  return (
    <Router>
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
        <FormModal isOpen={isFormModalOpen} onClose={handleCloseEnquiry} />
        <Toaster />
      </div>
    </Router>
  );
}

export default App;
