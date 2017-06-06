'use strict';

const express = require('express');
const router = express.Router();
const bodyParser = require('body-parser')

const development = process.env.NODE_ENV !== 'production';

var proofsDict = {}

router.use(bodyParser.json());

router.post('/stamp', function (req, res) {
  try {
    web3.setProvider(new Web3.providers.HttpProvider('http://10.0.0.4:8545'));
    var coinbase = web3.eth.coinbase;
    var originalBalance = web3.eth.getBalance(coinbase).toNumber();
    next(web3.eth.getBalance(coinbase));
    res.send({result: originalBalance, error: null});
  }
  catch (ex){
      res.status(503).send({error: ex.message});
  }
});

module.exports = router;
