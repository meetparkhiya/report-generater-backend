// routes/chatRoutes.js
const express = require("express");
const router = express.Router();

const {
    paginateChats,
} = require("../controllers/chatController");

router.post("/paginate", paginateChats);

module.exports = router;
