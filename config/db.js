// config/db.js
const mongoose = require("mongoose");

const connectDB = async () => {
    try {
        await mongoose.connect(process.env.MONGO_URI, {
            autoIndex: true,
        });
        console.log("MongoDB Connected ✅");
    } catch (err) {
        console.error("DB Error:", err.message);
        process.exit(1);
    }
};

module.exports = connectDB;
