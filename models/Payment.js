const mongoose = require("mongoose");

const paymentSchema =
new mongoose.Schema({

  registrationNo:String,

  amount:Number,

  paymentDate:Date,

  transactionId:String

},{
  timestamps:true
});

module.exports =
mongoose.model(
  "Payment",
  paymentSchema
);