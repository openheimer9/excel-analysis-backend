const express = require('express');
const router = express.Router();
const User = require('../models/User');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

// Middleware to verify JWT token
const authMiddleware = async (req, res, next) => {
  try {
    const token = req.header('Authorization').replace('Bearer ', '');
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    const user = await User.findById(decoded.id);
    
    if (!user) {
      throw new Error();
    }
    
    req.user = user;
    req.token = token;
    next();
  } catch (error) {
    res.status(401).send({ error: 'Please authenticate' });
  }
};

// Get user profile
router.get('/profile', authMiddleware, async (req, res) => {
  res.send(req.user);
});

// Get all users (admin only)
router.get('/users', authMiddleware, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).send({ error: 'Access denied' });
    }
    
    const users = await User.find({});
    res.send(users);
  } catch (error) {
    res.status(500).send({ error: 'Server error' });
  }
});

module.exports = router;
