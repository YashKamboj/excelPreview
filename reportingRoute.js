const express = require('express');

// Create an Express router
const router = express.Router();

// Middleware for token authentication
const authenticateToken = require('../middlewares/jwtAuth');

// Controllers for handling  fetching data of user assignments report 
const reportingController = require('../controllers/reportingController'); 

// Protected routes (require token authentication)
router.use(authenticateToken.authenticateToken);

// Route to get tables data
router.post('/getReport',reportingController.getReport);

// Export the router for use in the main application
module.exports = router;