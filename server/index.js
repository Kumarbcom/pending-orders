import express from 'express';
import cors from 'cors';
import mongoose from 'mongoose';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

// Middleware
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  methods: ['GET', 'POST', 'DELETE'],
  allowedHeaders: ['Content-Type']
}));
app.use(express.json({ limit: '10mb' }));

// MongoDB Connection
const MONGO_URI = process.env.MONGO_URI;
if (!MONGO_URI) {
  console.error('ERROR: MONGO_URI is not set in .env');
  process.exit(1);
}

mongoose.connect(MONGO_URI)
  .then(() => console.log('✅ MongoDB Atlas connected'))
  .catch(err => { console.error('❌ MongoDB connection error:', err); process.exit(1); });

// Generic schema for all data collections
const DataSchema = new mongoose.Schema({
  type:      { type: String, required: true, unique: true },
  data:      { type: mongoose.Schema.Types.Mixed, default: [] },
  updatedAt: { type: Date, default: Date.now }
});
const DataStore = mongoose.model('DataStore', DataSchema);

// Health check
app.get('/', (req, res) => res.json({ status: 'ok', message: 'Pending Orders API running' }));

// GET /api/:type — load a dataset
app.get('/api/:type', async (req, res) => {
  try {
    const doc = await DataStore.findOne({ type: req.params.type });
    res.json({ data: doc ? doc.data : [] });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST /api/:type — save / overwrite a dataset
app.post('/api/:type', async (req, res) => {
  try {
    const { data } = req.body;
    await DataStore.findOneAndUpdate(
      { type: req.params.type },
      { data, updatedAt: new Date() },
      { upsert: true, new: true }
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE /api/:type — clear a dataset
app.delete('/api/:type', async (req, res) => {
  try {
    await DataStore.findOneAndUpdate(
      { type: req.params.type },
      { data: [], updatedAt: new Date() },
      { upsert: true }
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
