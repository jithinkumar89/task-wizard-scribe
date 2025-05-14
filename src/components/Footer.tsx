
import React from 'react';

const Footer = () => {
  return (
    <footer className="py-3 mt-8 text-center text-gray-500 text-sm border-t">
      <div className="container mx-auto">
        <p>© {new Date().getFullYear()} SOP Processor | V 1.0</p>
      </div>
    </footer>
  );
};

export default Footer;
