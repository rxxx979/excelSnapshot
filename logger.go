package excelsnapshot

import (
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

// SetupLogger 设置日志
func SetupLogger(name string, level zapcore.Level, isDev bool) (*zap.Logger, func(), error) {
	var cfg zap.Config
	if isDev {
		cfg = zap.NewDevelopmentConfig()
	} else {
		cfg = zap.NewProductionConfig()
	}
	cfg.Level = zap.NewAtomicLevelAt(level)

	logger, err := cfg.Build()
	if err != nil {
		return nil, func() {}, err
	}

	logger = logger.Named(name)
	return logger, func() { _ = logger.Sync() }, nil
}
