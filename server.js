const express = require("express");
const path = require("path");
const session = require("express-session");
const expressLayouts = require("express-ejs-layouts");

const app = express();

// View engine
app.set("view engine", "ejs");
app.use(expressLayouts);
app.set("layout", "layout");
app.set("views", path.join(__dirname, "views"));

// Static files & session
app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: "secretKey",
  resave: false,
  saveUninitialized: false
}));

// Auth middleware
function requireLogin(req, res, next) {
  if (!req.session.user) return res.redirect("/login");
  next();
}

// Global user
app.use((req, res, next) => {
  res.locals.user = req.session.user;
  next();
});

// Routes
app.get("/", (req, res) => res.redirect("/login"));
app.get("/login", (req, res) => res.render("login", { layout: false, error: null }));

app.post("/login", (req, res) => {
  const { username, password } = req.body;
  if (username === "admin" && password === "1234") {
    req.session.user = { username };
    return res.redirect("/dashboard");
  }
  res.render("login", { layout: false, error: "Invalid credentials!" });
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});

// Dashboard
app.get("/dashboard", requireLogin, (req, res) => res.render("dashboard", { page: "dashboard" }));

// Rollingan
app.get("/rollingan/cashback", requireLogin, (req, res) => res.render("rollingan_cashback", { page: "cashback" }));
app.get("/rollingan/noncashback", requireLogin, (req, res) => res.render("rollingan_noncashback", { page: "noncashback" }));

// Transaction
app.get("/transaction/pivot", requireLogin, (req, res) => res.render("pivot", { page: "pivot" }));
app.get("/transaction/referral", requireLogin, (req, res) => res.render("referral", { page: "transaction_referral" }));

// Report
app.get("/report", requireLogin, (req, res) => res.render("report", { page: "report" }));

// Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running at http://localhost:${PORT}`));
