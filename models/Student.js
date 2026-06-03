const mongoose = require("mongoose");

const studentSchema = new mongoose.Schema({

  registrationNo:{
    type:String,
    unique:true
  },

  aadhaar:{
    type:String,
    unique:true
  },

  name:String,
  phone:String,
  email:String,

  college:String,
  course:String,
  branch:String,

  status:{
    type:String,
    default:"Pending"
  }

},{
  timestamps:true
});

studentSchema.index({
  registrationNo:1
});

studentSchema.index({
  aadhaar:1
});

module.exports =
mongoose.model(
  "Student",
  studentSchema
);