# 🔥 LPG Gas Management System

A desktop-based application developed using **Visual Basic 6.0** and **Microsoft Access** to manage complete LPG (Liquefied Petroleum Gas) distribution operations.

---

## 📌 About The Project

The LPG Gas Management System is designed to replace manual paper-based record keeping in LPG distribution agencies. It provides a fast, accurate, and user-friendly solution for managing customers, bookings, billing, and stock inventory.

---

## 🛠️ Built With

| Technology | Purpose |
|------------|---------|
| Visual Basic 6.0 | Front-end Development |
| Microsoft Access 2019 | Database Management |
| ADO Data Control 6.0 | Database Connectivity |
| Microsoft DataGrid Control | Data Display |
| ACE OLEDB 12.0 | Database Provider |

---

## 📋 Features

- ✅ **Splash Screen** - Animated loading screen with progress bar
- ✅ **Secure Login** - Username and password authentication
- ✅ **Dashboard** - Central navigation with real-time clock
- ✅ **Customer Registration** - Full CRUD operations for customer management
- ✅ **Gas Booking System** - Manage cylinder bookings with status tracking
- ✅ **Bill Generation** - Generate and manage bills with payment status
- ✅ **Stock Management** - Track cylinder inventory for multiple types

---

## 🗄️ Database Tables

| Table | Description |
|-------|-------------|
| tblCustomer | Stores customer registration details |
| tblBooking | Manages gas cylinder booking records |
| tblBill | Handles billing and payment records |
| tblStock | Tracks cylinder inventory |

---

## 📁 Project Structure

```
LPG-Gas-Management-System/
│
├── frmSplash.frm       # Splash Screen Form
├── frmLogin.frm        # Login Form
├── frmDashboard.frm    # Dashboard Form
├── frmCustomer.frm     # Customer Registration Form
├── frmCustomer.frx     # Customer Form Binary
├── frmBooking.frm      # Gas Booking Form
├── frmBooking.frx      # Booking Form Binary
├── frmBill.frm         # Bill Generation Form
├── frmBill.frx         # Bill Form Binary
├── frmStock.frm        # Stock Management Form
├── frmStock.frx        # Stock Form Binary
└── LPGGasSystem.vbp    # VB6 Project File
```

---

## 🚀 How To Run

### Requirements
- Windows XP / 7 / 10 / 11
- Visual Basic 6.0 Enterprise Edition
- Microsoft Access 2007 or higher
- Microsoft ACE OLEDB 12.0 Provider

### Steps
1. Clone or download this repository
2. Open `LPGGasSystem.vbp` in Visual Basic 6.0
3. Create database `LPGGasSystem.accdb` in Microsoft Access
4. Update the database path in each form's `Form_Load` event
5. Press `F5` to run the project

### Login Credentials
```
Username : admin
Password : admin123
```

---

## 📸 Screenshots

### Splash Screen
> Animated loading screen with progress bar

### Login Form
> Secure login with username and password

### Dashboard
> Main navigation hub with 4 module buttons and real-time clock

### Customer Registration
> Full CRUD operations with live DataGrid display

### Gas Booking System
> Manage bookings with status tracking (Pending/Processing/Delivered)

### Bill Generation
> Generate bills linked to customers and bookings

### Stock Management
> Track cylinder inventory for 5KG, 14.2KG, 19KG, 47.5KG types

---

## 👨‍💻 Developer

| Detail | Info |
|--------|------|
| **Name** | Neelkamal Jangde |
| **Course** | PGDCA (Post Graduate Diploma in Computer Applications) |
| **University** | ISBM University |
| **Year** | 2025 - 2026 |

---

## 📄 License

This project is developed for academic purposes as part of PGDCA Final Year Project at ISBM University.

---

> *"Your Trusted Gas Partner"* 🔥
