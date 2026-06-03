const mongoose =
require("mongoose");

const logSchema =
new mongoose.Schema({

  action:String,

  admin:String,

  time:{
    type:Date,
    default:Date.now
  }

});


module.exports =
mongoose.model(
  "ActivityLog",
  logSchema
);
const adminSchema = new mongoose.Schema({
  email: {
    type: String,
    unique: true
  },
  phone: {
    type: String,
    unique: true
  },
  password: String,
  resetOtp: String,
  resetOtpExpire: Date
}, { timestamps: true });

const Admin = mongoose.model("Admin", adminSchema);

