const Chat = require("../models/Chat");

const PAGE_SIZE = 5;

exports.paginateChats = async (req, res) => {
    try {
        const { excludeIds = [], per_page = PAGE_SIZE, search = "" } = req.body;

        const query = {};

        // filter already loaded IDs
        if (excludeIds.length > 0) {
            query._id = { $nin: excludeIds };
        }

        // search filter
        if (search) {
            query.name = { $regex: search, $options: "i" };
        }

        // total in collection (all records, ignore filter)
        const totalInDB = await Chat.estimatedDocumentCount();

        // total matching current filter
        const totalMatching = await Chat.countDocuments(query);

        const chats = await Chat.find(query)
            .sort({ createdAt: 1 })  // oldest first
            .limit(per_page);

        res.json({
            data: chats,
            totalInDB,         // total chat records in DB
            totalMatching,     // total matching records for filter
            hasMore: chats.length === per_page,
        });
    } catch (err) {
        console.error("Error Fetch Chat List:", err.message);
        res.status(500).json({ message: "Server Error", error: err.message });
    }
};
