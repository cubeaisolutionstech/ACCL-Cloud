// components/SearchableSelect.jsx
import React, { useState, useEffect, useRef } from 'react';

const SearchableSelect = ({
  options = [],
  value = '',
  onChange,
  placeholder = 'Select...',
  className = 'w-full p-2 border',
  disabled = false,
  emptyText = 'No options found'
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filteredOptions, setFilteredOptions] = useState(options);
  const [highlightedIndex, setHighlightedIndex] = useState(-1);
  
  const containerRef = useRef(null);
  const inputRef = useRef(null);
  const optionsRef = useRef([]);

  // Filter options based on search term
  useEffect(() => {
    if (searchTerm.trim() === '') {
      setFilteredOptions(options);
    } else {
      const filtered = options.filter(option =>
        option.toLowerCase().includes(searchTerm.toLowerCase())
      );
      setFilteredOptions(filtered);
    }
    setHighlightedIndex(-1);
  }, [searchTerm, options]);

  // Update search term when value changes externally
  useEffect(() => {
    if (!isOpen) {
      setSearchTerm(value);
    }
  }, [value, isOpen]);

  // Handle click outside to close dropdown
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (containerRef.current && !containerRef.current.contains(event.target)) {
        setIsOpen(false);
        setSearchTerm(value); // Reset to selected value
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [value]);

  // Handle keyboard navigation
  const handleKeyDown = (e) => {
    if (disabled) return;

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        setIsOpen(true);
        setHighlightedIndex(prev => 
          prev < filteredOptions.length - 1 ? prev + 1 : 0
        );
        break;
      
      case 'ArrowUp':
        e.preventDefault();
        if (isOpen) {
          setHighlightedIndex(prev => 
            prev > 0 ? prev - 1 : filteredOptions.length - 1
          );
        }
        break;
      
      case 'Enter':
        e.preventDefault();
        if (isOpen && highlightedIndex >= 0 && filteredOptions[highlightedIndex]) {
          handleOptionSelect(filteredOptions[highlightedIndex]);
        } else if (!isOpen) {
          setIsOpen(true);
        }
        break;
      
      case 'Escape':
        e.preventDefault();
        setIsOpen(false);
        setSearchTerm(value);
        setHighlightedIndex(-1);
        break;
      
      case 'Tab':
        setIsOpen(false);
        setSearchTerm(value);
        break;
      
      default:
        if (!isOpen) {
          setIsOpen(true);
        }
        break;
    }
  };

  // Handle option selection
  const handleOptionSelect = (option) => {
    onChange(option);
    setSearchTerm(option);
    setIsOpen(false);
    setHighlightedIndex(-1);
    inputRef.current?.blur();
  };

  // Handle input focus
  const handleFocus = () => {
    if (!disabled) {
      setIsOpen(true);
      setSearchTerm(''); // Clear search to show all options
    }
  };

  // Handle input change
  const handleInputChange = (e) => {
    setSearchTerm(e.target.value);
    if (!isOpen) {
      setIsOpen(true);
    }
  };

  // Scroll highlighted option into view
  useEffect(() => {
    if (highlightedIndex >= 0 && optionsRef.current[highlightedIndex]) {
      optionsRef.current[highlightedIndex].scrollIntoView({
        block: 'nearest',
        behavior: 'smooth'
      });
    }
  }, [highlightedIndex]);

  return (
    <div className="relative" ref={containerRef}>
      {/* Input field */}
      <input
        ref={inputRef}
        type="text"
        className={`${className} ${disabled ? 'bg-gray-100 cursor-not-allowed' : 'cursor-pointer'}`}
        value={searchTerm}
        onChange={handleInputChange}
        onFocus={handleFocus}
        onKeyDown={handleKeyDown}
        placeholder={value || placeholder}
        disabled={disabled}
        autoComplete="off"
        style={{ 
          paddingRight: '2rem',
          backgroundColor: disabled ? '#f3f4f6' : '#ffffff'
        }}
      />
      
      {/* Dropdown arrow */}
      <div 
        className="absolute right-2 top-1/2 transform -translate-y-1/2 pointer-events-none"
        style={{ color: disabled ? '#9ca3af' : '#6b7280' }}
      >
        <svg 
          className={`w-4 h-4 transition-transform ${isOpen ? 'rotate-180' : ''}`} 
          fill="none" 
          stroke="currentColor" 
          viewBox="0 0 24 24"
        >
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </div>

      {/* Dropdown options */}
      {isOpen && !disabled && (
        <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-md shadow-lg max-h-60 overflow-y-auto">
          {filteredOptions.length > 0 ? (
            filteredOptions.map((option, index) => (
              <div
                key={`${option}-${index}`}
                ref={el => optionsRef.current[index] = el}
                className={`px-3 py-2 cursor-pointer hover:bg-blue-100 ${
                  index === highlightedIndex 
                    ? 'bg-blue-100 text-blue-800' 
                    : option === value 
                      ? 'bg-blue-50 text-blue-700 font-medium' 
                      : 'text-gray-800'
                }`}
                onClick={() => handleOptionSelect(option)}
                onMouseEnter={() => setHighlightedIndex(index)}
              >
                {option}
              </div>
            ))
          ) : (
            <div className="px-3 py-2 text-gray-500 italic">
              {emptyText}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default SearchableSelect;