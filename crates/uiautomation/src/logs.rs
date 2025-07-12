#[macro_export]
macro_rules! log_debug {
    ($($arg:tt)*) => {
        #[cfg(feature = "log")]
        log::debug!($($arg)*);
        #[cfg(not(feature = "log"))]
        println!($($arg)*);
    };
}

#[macro_export]
macro_rules! log_error {
    ($($arg:tt)*) => {
        #[cfg(feature = "log")]
        log::error!($($arg)*);
        #[cfg(not(feature = "log"))]
        eprintln!($($arg)*);
    };
}
