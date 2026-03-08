# Header Styling Guide - CSS Structure for Green-Themed Project

## Overview
This guide provides the exact CSS structure used for headers in both main and training pages, adapted for a green color scheme. All font sizes, spacing, and responsive breakpoints are preserved.

## Color Scheme Conversion

### Original Colors:
- **Main Site**: Blue (`from-blue-600 to-blue-700`, `text-blue-100`)
- **Training Site**: Orange (`from-orange-600 to-orange-700`, `text-orange-100`)

### New Green Colors:
- **Primary Green Header**: `from-green-600 to-green-700`
- **Secondary Green Text**: `text-green-100`
- **Icon Backgrounds**: `bg-green-100`, `text-green-600`
- **Hover States**: `hover:bg-green-50`

## Main Header Component Structure

### Core Header Layout
```tsx
<header className="bg-gradient-to-r from-green-600 to-green-700 text-white shadow-lg fixed top-0 left-0 right-0 z-30">
  <div className="container mx-auto px-4">
    <div className="flex items-center justify-between h-16">
      {/* Content */}
    </div>
  </div>
</header>
```

### Logo Section with Site Selector
```tsx
<div className="relative">
  <div className="flex items-center space-x-2 sm:space-x-3">
    <div className="flex items-center space-x-2 sm:space-x-3">
      {/* Site Selector Icon */}
      <div className="w-8 h-8 sm:w-10 sm:h-10 bg-white bg-opacity-20 rounded-lg flex items-center justify-center flex-shrink-0 cursor-pointer hover:bg-opacity-30 transition-all">
        <Shield className="w-4 h-4 sm:w-5 sm:h-5 text-white" />
      </div>
      
      {/* Company Name and Subtitle */}
      <Link href="/">
        <div className="min-w-0 flex-1 pr-2 cursor-pointer hover:opacity-90 transition-opacity">
          <h1 className="text-base sm:text-xl font-bold text-white truncate">Your Company Name</h1>
          <p className="text-green-100 text-xs sm:text-sm truncate">Financial Pricing System</p>
        </div>
      </Link>
    </div>
  </div>
</div>
```

### Typography Specifications
- **Main Title**: `text-base sm:text-xl font-bold text-white truncate`
- **Subtitle**: `text-green-100 text-xs sm:text-sm truncate`
- **Mobile breakpoint**: `sm:` prefix for 640px+
- **Font weight**: `font-bold` for titles, normal for subtitles

### Site Selector Dropdown
```tsx
{isSiteDropdownOpen && (
  <div className="absolute top-full left-0 mt-2 w-72 bg-white rounded-lg shadow-xl border z-[70]">
    <div className="p-4">
      <h3 className="text-sm font-medium text-gray-900 mb-3">Available Sites</h3>
      <div className="space-y-2">
        {/* Financial Site */}
        <Link href="/">
          <div className="flex items-center space-x-3 p-3 rounded-lg hover:bg-green-50 cursor-pointer transition-colors">
            <div className="w-8 h-8 bg-green-100 rounded-lg flex items-center justify-center">
              <DollarSign className="w-4 h-4 text-green-600" />
            </div>
            <div>
              <div className="text-sm font-medium text-gray-900">Financial Pricing System</div>
              <div className="text-xs text-gray-500">Pricing & Cost Management</div>
            </div>
          </div>
        </Link>
        
        {/* Additional Sites */}
        <Link href="/other-section">
          <div className="flex items-center space-x-3 p-3 rounded-lg hover:bg-blue-50 cursor-pointer transition-colors">
            <div className="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center">
              <FileText className="w-4 h-4 text-blue-600" />
            </div>
            <div>
              <div className="text-sm font-medium text-gray-900">Other Section</div>
              <div className="text-xs text-gray-500">Description here</div>
            </div>
          </div>
        </Link>
      </div>
    </div>
  </div>
)}
```

## Navigation Buttons Structure

### Desktop Navigation
```tsx
<div className="hidden md:flex items-center space-x-2 lg:space-x-4">
  <Link href="/pricing-history">
    <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-2 lg:px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-1 lg:space-x-2">
      <Calculator className="w-4 h-4 flex-shrink-0" />
      <span className="hidden lg:inline">Pricing History</span>
      <span className="lg:hidden">Pricing History</span>
    </button>
  </Link>
</div>
```

### Tablet Navigation (Hidden on mobile, shown between sm and md)
```tsx
<div className="hidden sm:flex md:hidden items-center space-x-2">
  <Link href="/pricing-history">
    <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white w-9 h-9 rounded-lg transition-all duration-200 flex items-center justify-center" title="Pricing History">
      <Calculator className="w-4 h-4" />
    </button>
  </Link>
</div>
```

### Authentication Buttons
```tsx
{!isAuthenticated ? (
  <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-2 lg:px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-1 lg:space-x-2 disabled:opacity-50">
    <User className="w-4 h-4 flex-shrink-0" />
    <span className="hidden lg:inline">Sign In</span>
    <span className="lg:hidden">Login</span>
  </button>
) : (
  <div className="relative">
    <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-2 lg:px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-1 lg:space-x-2 h-9 min-w-0">
      <User className="w-4 h-4 flex-shrink-0" />
      <span className="truncate text-xs lg:text-sm">{userInfo?.name || 'User'}</span>
      <ChevronDown className="w-3 h-3 lg:w-4 lg:h-4 flex-shrink-0" />
    </button>
  </div>
)}
```

## Responsive Breakpoints

### Screen Size Classes
- **Mobile**: Default (320px+) - Hidden navigation, hamburger menu
- **Small**: `sm:` (640px+) - Show tablet navigation, hide desktop text
- **Medium**: `md:` (768px+) - Show full desktop navigation
- **Large**: `lg:` (1024px+) - Show full button text and spacing

### Button Sizing
- **Mobile**: `w-9 h-9` (square buttons, icons only)
- **Desktop**: `px-2 lg:px-4 py-2` (rectangular with text)
- **Icon sizes**: `w-4 h-4` for buttons, `w-4 h-4 sm:w-5 sm:h-5` for logo

## CSS Variables for Green Theme

### Add to your CSS file:
```css
:root {
  /* Green theme colors */
  --green-50: hsl(142, 76%, 96%);
  --green-100: hsl(141, 84%, 93%);
  --green-600: hsl(142, 71%, 45%);
  --green-700: hsl(142, 72%, 38%);
  
  /* Keep existing Microsoft colors for multi-site support */
  --ms-blue: hsl(207, 100%, 41%);
  --ms-orange: hsl(25, 95%, 53%);
  --ms-green: hsl(142, 71%, 45%); /* Your new primary */
}

/* Green theme classes */
.bg-green-600 {
  background-color: var(--ms-green);
}

.bg-green-700 {
  background-color: hsl(142, 72%, 38%);
}

.text-green-100 {
  color: hsl(141, 84%, 93%);
}

.text-green-600 {
  color: var(--ms-green);
}

.bg-green-100 {
  background-color: hsl(141, 84%, 93%);
}

.bg-green-50 {
  background-color: hsl(142, 76%, 96%);
}

.hover\:bg-green-50:hover {
  background-color: hsl(142, 76%, 96%);
}
```

## Complete Green Header Example

```tsx
import { Link } from "wouter";
import { useState } from "react";
import { Shield, Calculator, User, ChevronDown, DollarSign, FileText } from "lucide-react";

export default function GreenHeader() {
  const [isSiteDropdownOpen, setIsSiteDropdownOpen] = useState(false);

  return (
    <header className="bg-gradient-to-r from-green-600 to-green-700 text-white shadow-lg fixed top-0 left-0 right-0 z-30">
      <div className="container mx-auto px-4">
        <div className="flex items-center justify-between h-16">
          {/* Logo with Site Selector */}
          <div className="relative">
            <div className="flex items-center space-x-2 sm:space-x-3">
              <div className="flex items-center space-x-2 sm:space-x-3">
                <div 
                  className="w-8 h-8 sm:w-10 sm:h-10 bg-white bg-opacity-20 rounded-lg flex items-center justify-center flex-shrink-0 cursor-pointer hover:bg-opacity-30 transition-all"
                  onClick={() => setIsSiteDropdownOpen(!isSiteDropdownOpen)}
                  title="Site selector"
                >
                  <Shield className="w-4 h-4 sm:w-5 sm:h-5 text-white" />
                </div>
                <Link href="/">
                  <div className="min-w-0 flex-1 pr-2 cursor-pointer hover:opacity-90 transition-opacity">
                    <h1 className="text-base sm:text-xl font-bold text-white truncate">Your Company Name</h1>
                    <p className="text-green-100 text-xs sm:text-sm truncate">Financial Pricing System</p>
                  </div>
                </Link>
              </div>
            </div>
            
            {/* Site Dropdown */}
            {isSiteDropdownOpen && (
              <div className="absolute top-full left-0 mt-2 w-72 bg-white rounded-lg shadow-xl border z-[70]">
                <div className="p-4">
                  <h3 className="text-sm font-medium text-gray-900 mb-3">Available Sites</h3>
                  <div className="space-y-2">
                    <Link href="/">
                      <div className="flex items-center space-x-3 p-3 rounded-lg hover:bg-green-50 cursor-pointer transition-colors">
                        <div className="w-8 h-8 bg-green-100 rounded-lg flex items-center justify-center">
                          <DollarSign className="w-4 h-4 text-green-600" />
                        </div>
                        <div>
                          <div className="text-sm font-medium text-gray-900">Financial Pricing System</div>
                          <div className="text-xs text-gray-500">Pricing & Cost Management</div>
                        </div>
                      </div>
                    </Link>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Desktop Navigation */}
          <div className="hidden md:flex items-center space-x-2 lg:space-x-4">
            <Link href="/pricing-history">
              <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-2 lg:px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-1 lg:space-x-2">
                <Calculator className="w-4 h-4 flex-shrink-0" />
                <span className="hidden lg:inline">Pricing History</span>
                <span className="lg:hidden">Pricing History</span>
              </button>
            </Link>

            {/* Authentication button would go here */}
          </div>

          {/* Tablet Navigation */}
          <div className="hidden sm:flex md:hidden items-center space-x-2">
            <Link href="/pricing-history">
              <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white w-9 h-9 rounded-lg transition-all duration-200 flex items-center justify-center" title="Pricing History">
                <Calculator className="w-4 h-4" />
              </button>
            </Link>
          </div>
        </div>
      </div>

      {/* Dropdown backdrop */}
      {isSiteDropdownOpen && (
        <div 
          className="fixed inset-0 z-40" 
          onClick={() => setIsSiteDropdownOpen(false)}
        />
      )}
    </header>
  );
}
```

## Key Styling Features

### Professional Appearance
- **Fixed header**: `fixed top-0 left-0 right-0 z-30`
- **Gradient background**: `bg-gradient-to-r from-green-600 to-green-700`
- **Shadow**: `shadow-lg`
- **Height**: `h-16` (64px consistent)

### Interactive Elements
- **Hover states**: `hover:bg-opacity-30`, `hover:opacity-90`
- **Transitions**: `transition-all duration-200`
- **Rounded corners**: `rounded-lg` (8px)
- **Semi-transparent backgrounds**: `bg-white bg-opacity-20`

### Typography
- **Font family**: Uses Inter font from CSS (`font-inter`)
- **Title**: `font-bold` with responsive sizing
- **Buttons**: `text-sm font-medium`
- **Truncation**: `truncate` for long text

### Mobile Responsiveness
- **Container**: `container mx-auto px-4`
- **Spacing**: `space-x-2 sm:space-x-3`
- **Icon sizing**: `w-4 h-4 sm:w-5 sm:h-5`
- **Text visibility**: `hidden lg:inline` for responsive text

This structure provides the exact same professional appearance and functionality as the original headers, just with a green color scheme perfect for a financial pricing system!