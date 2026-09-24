import React from "react";

const menuItems = [
  { id: "start", label: "START HERE" },
  { id: "books", label: "BOOKS INPUT" },
  { id: "2b", label: "2B INPUT" },
  { id: "recon", label: "RECON" },
  { id: "summary", label: "SUMMARY" },
  { id: "matched", label: "MATCHED" },
  { id: "duplicate", label: "DUPLICATE" },
  { id: "not-books", label: "NOT IN BOOKS" },
  { id: "not-2b", label: "NOT IN 2B" },
  { id: "rcm", label: "RCM" },
  { id: "imports", label: "IMPORTS" },
  { id: "credit-notes", label: "DR CREDIT NOTES" },
];

function Sidebar({ activeSection, onSectionChange }) {
  return (
    <aside className="w-64 min-h-screen bg-white border-r border-gray-200 flex flex-col">

      {/* Logo / Brand */}
      <div className="px-6 py-6 border-b border-gray-200">
        <div className="text-2xl font-bold text-green-700">
          GST RECO
        </div>

        <div className="text-xs text-gray-500 mt-1">
          Smart GST Reconciliation
        </div>
      </div>

      {/* Navigation */}
      <nav className="flex-1 px-3 py-5">

        {menuItems.map((item, index) => {
          const isActive = activeSection === item.id;

          return (
            <button
              key={item.id}
              onClick={() => onSectionChange(item.id)}
              className={`w-full flex items-center gap-3 px-4 py-3 mb-1 rounded-lg text-left transition-all duration-200 ${
                isActive
                  ? "bg-green-100 text-green-700 font-semibold"
                  : "text-gray-700 hover:bg-gray-100"
              }`}
            >

              {/* Number */}
              <span
                className={`w-6 text-sm font-bold ${
                  isActive ? "text-green-700" : "text-gray-500"
                }`}
              >
                {index + 1}.
              </span>

              {/* Menu name */}
              <span className="text-sm">
                {item.label}
              </span>

            </button>
          );
        })}

      </nav>

      {/* Bottom section */}
      <div className="px-6 py-5 border-t border-gray-200">
        <div className="text-xs text-gray-400">
          GST RECO
        </div>

        <div className="text-xs text-gray-400 mt-1">
          GST Reconciliation System
        </div>
      </div>

    </aside>
  );
}

export default Sidebar;