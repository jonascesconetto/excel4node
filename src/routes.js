const express = require("express");
const OrderController = require('./controllers/OrderController');

const routes = express.Router();

routes.get('/', (req, res) => {
    return res.send('Hello World')
});

routes.post('/orders', OrderController.store); // criação do arquivo

module.exports = routes;
