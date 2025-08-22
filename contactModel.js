import mongoose from 'mongoose';

const contactSchema = new mongoose.Schema({
    fullName: { type: String, required: true },
    email: { type: String, required: true },
    phone: { type: String },
    subject: { type: String, required: true },
    message: { type: String, required: true },
    preferredContact: { type: String, required: true, enum: ["email", "phone"] },
  }, { timestamps: true, collection: "Form" });
  
export default mongoose.model("Contact", contactSchema);
  