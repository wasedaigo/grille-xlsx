"use strict";

const winston = require('winston');

// logger.
// https://github.com/winstonjs/winston#using-the-default-logger
module.exports = new winston.Logger(
    {
        transports: [
            new winston.transports.Console(
                {
                    level: 'debug',
                    timestamp: true,
                    handleExceptions: true,
                    humanReadableUnhandledException: true,
                    json: false,
                    colorize: true
                }
            )
        ],
        exitOnError: true
    }
);
