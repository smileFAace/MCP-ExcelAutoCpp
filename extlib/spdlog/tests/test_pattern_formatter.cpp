#include "includes.h"
#include "test_sink.h"

#include <chrono>

using spdlog::memory_buf_t;
using spdlog::details::to_string_view;

// log to str and return it
template <typename... Args>
static std::string log_to_str(const std::string &msg, const Args &...args) {
    std::ostringstream oss;
    auto oss_sink = std::make_shared<spdlog::sinks::ostream_sink_mt>(oss);
    spdlog::logger oss_logger("pattern_tester", oss_sink);
    oss_logger.set_level(spdlog::level::info);

    oss_logger.set_formatter(
        std::unique_ptr<spdlog::formatter>(new spdlog::pattern_formatter(args...)));

    oss_logger.info(msg);
    return oss.str();
}

// log to str and return it with time
template <typename... Args>
static std::string log_to_str_with_time(spdlog::log_clock::time_point log_time,
                                        const std::string &msg,
                                        const Args &...args) {
    std::ostringstream oss;
    auto oss_sink = std::make_shared<spdlog::sinks::ostream_sink_mt>(oss);
    spdlog::logger oss_logger("pattern_tester", oss_sink);
    oss_logger.set_level(spdlog::level::info);

    oss_logger.set_formatter(
        std::unique_ptr<spdlog::formatter>(new spdlog::pattern_formatter(args...)));

    oss_logger.log(log_time, {}, spdlog::level::info, msg);
    return oss.str();
}

TEST_CASE("custom eol", "[pattern_formatter]") {
    std::string msg = "Hello custom eol test";
    std::string eol = ";)";
    REQUIRE(log_to_str(msg, "%v", spdlog::pattern_time_type::local, ";)") == msg + eol);
}

TEST_CASE("empty format", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "", spdlog::pattern_time_type::local, "").empty());
}

TEST_CASE("empty format2", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "", spdlog::pattern_time_type::local, "\n") == "\n");
}

TEST_CASE("level", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[info] Some message\n");
}

TEST_CASE("short level", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[I] Some message\n");
}

TEST_CASE("name", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester] Some message\n");
}

TEST_CASE("date MM/DD/YY ", "[pattern_formatter]") {
    auto now_tm = spdlog::details::os::localtime();
    std::stringstream oss;
    oss << std::setfill('0') << std::setw(2) << now_tm.tm_mon + 1 << "/" << std::setw(2)
        << now_tm.tm_mday << "/" << std::setw(2) << (now_tm.tm_year + 1900) % 1000
        << " Some message\n";
    REQUIRE(log_to_str("Some message", "%D %v", spdlog::pattern_time_type::local, "\n") ==
            oss.str());
}

TEST_CASE("GMT offset ", "[pattern_formatter]") {
    using namespace std::chrono_literals;
    const auto now = std::chrono::system_clock::now();
    const auto yesterday = now - 24h;

    REQUIRE(log_to_str_with_time(yesterday, "Some message", "%z", spdlog::pattern_time_type::utc,
                                 "\n") == "+00:00\n");
}

TEST_CASE("color range test1", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>(
        "%^%v%$", spdlog::pattern_time_type::local, "\n");

    memory_buf_t buf;
    spdlog::fmt_lib::format_to(std::back_inserter(buf), "Hello");
    memory_buf_t formatted;
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info,
                                 spdlog::string_view_t(buf.data(), buf.size()));
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 0);
    REQUIRE(msg.color_range_end == 5);
    REQUIRE(log_to_str("hello", "%^%v%$", spdlog::pattern_time_type::local, "\n") == "hello\n");
}

TEST_CASE("color range test2", "[pattern_formatter]") {
    auto formatter =
        std::make_shared<spdlog::pattern_formatter>("%^%$", spdlog::pattern_time_type::local, "\n");
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "");
    memory_buf_t formatted;
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 0);
    REQUIRE(msg.color_range_end == 0);
    REQUIRE(log_to_str("", "%^%$", spdlog::pattern_time_type::local, "\n") == "\n");
}

TEST_CASE("color range test3", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>("%^***%$");
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "ignored");
    memory_buf_t formatted;
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 0);
    REQUIRE(msg.color_range_end == 3);
}

TEST_CASE("color range test4", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>(
        "XX%^YYY%$", spdlog::pattern_time_type::local, "\n");
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "ignored");

    memory_buf_t formatted;
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 2);
    REQUIRE(msg.color_range_end == 5);
    REQUIRE(log_to_str("ignored", "XX%^YYY%$", spdlog::pattern_time_type::local, "\n") ==
            "XXYYY\n");
}

TEST_CASE("color range test5", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>("**%^");
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "ignored");
    memory_buf_t formatted;
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 2);
    REQUIRE(msg.color_range_end == 0);
}

TEST_CASE("color range test6", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>("**%$");
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "ignored");
    memory_buf_t formatted;
    formatter->format(msg, formatted);
    REQUIRE(msg.color_range_start == 0);
    REQUIRE(msg.color_range_end == 2);
}

//
// Test padding
//

TEST_CASE("level_left_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%8l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[    info] Some message\n");
    REQUIRE(log_to_str("Some message", "[%8!l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[    info] Some message\n");
}

TEST_CASE("level_right_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%-8l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[info    ] Some message\n");
    REQUIRE(log_to_str("Some message", "[%-8!l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[info    ] Some message\n");
}

TEST_CASE("level_center_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%=8l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[  info  ] Some message\n");
    REQUIRE(log_to_str("Some message", "[%=8!l] %v", spdlog::pattern_time_type::local, "\n") ==
            "[  info  ] Some message\n");
}

TEST_CASE("short level_left_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%3L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[  I] Some message\n");
    REQUIRE(log_to_str("Some message", "[%3!L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[  I] Some message\n");
}

TEST_CASE("short level_right_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%-3L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[I  ] Some message\n");
    REQUIRE(log_to_str("Some message", "[%-3!L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[I  ] Some message\n");
}

TEST_CASE("short level_center_padded", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%=3L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[ I ] Some message\n");
    REQUIRE(log_to_str("Some message", "[%=3!L] %v", spdlog::pattern_time_type::local, "\n") ==
            "[ I ] Some message\n");
}

TEST_CASE("left_padded_short", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%3n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester] Some message\n");
    REQUIRE(log_to_str("Some message", "[%3!n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pat] Some message\n");
}

TEST_CASE("right_padded_short", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%-3n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester] Some message\n");
    REQUIRE(log_to_str("Some message", "[%-3!n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pat] Some message\n");
}

TEST_CASE("center_padded_short", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%=3n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester] Some message\n");
    REQUIRE(log_to_str("Some message", "[%=3!n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pat] Some message\n");
}

TEST_CASE("left_padded_huge", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%-300n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester                                                  ] Some message\n");

    REQUIRE(log_to_str("Some message", "[%-300!n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester                                                  ] Some message\n");
}

TEST_CASE("left_padded_max", "[pattern_formatter]") {
    REQUIRE(log_to_str("Some message", "[%-64n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester                                                  ] Some message\n");

    REQUIRE(log_to_str("Some message", "[%-64!n] %v", spdlog::pattern_time_type::local, "\n") ==
            "[pattern_tester                                                  ] Some message\n");
}

// Test padding + truncate flag

TEST_CASE("paddinng_truncate", "[pattern_formatter]") {
    REQUIRE(log_to_str("123456", "%6!v", spdlog::pattern_time_type::local, "\n") == "123456\n");
    REQUIRE(log_to_str("123456", "%5!v", spdlog::pattern_time_type::local, "\n") == "12345\n");
    REQUIRE(log_to_str("123456", "%7!v", spdlog::pattern_time_type::local, "\n") == " 123456\n");

    REQUIRE(log_to_str("123456", "%-6!v", spdlog::pattern_time_type::local, "\n") == "123456\n");
    REQUIRE(log_to_str("123456", "%-5!v", spdlog::pattern_time_type::local, "\n") == "12345\n");
    REQUIRE(log_to_str("123456", "%-7!v", spdlog::pattern_time_type::local, "\n") == "123456 \n");

    REQUIRE(log_to_str("123456", "%=6!v", spdlog::pattern_time_type::local, "\n") == "123456\n");
    REQUIRE(log_to_str("123456", "%=5!v", spdlog::pattern_time_type::local, "\n") == "12345\n");
    REQUIRE(log_to_str("123456", "%=7!v", spdlog::pattern_time_type::local, "\n") == "123456 \n");

    REQUIRE(log_to_str("123456", "%0!v", spdlog::pattern_time_type::local, "\n") == "\n");
}

TEST_CASE("padding_truncate_funcname", "[pattern_formatter]") {
    spdlog::sinks::test_sink_st test_sink;

    const char *pattern = "%v [%5!!]";
    auto formatter = std::unique_ptr<spdlog::formatter>(new spdlog::pattern_formatter(pattern));
    test_sink.set_formatter(std::move(formatter));

    spdlog::details::log_msg msg1{spdlog::source_loc{"ignored", 1, "func"}, "test_logger",
                                  spdlog::level::info, "message"};
    test_sink.log(msg1);
    REQUIRE(test_sink.lines()[0] == "message [ func]");

    spdlog::details::log_msg msg2{spdlog::source_loc{"ignored", 1, "function"}, "test_logger",
                                  spdlog::level::info, "message"};
    test_sink.log(msg2);
    REQUIRE(test_sink.lines()[1] == "message [funct]");
}

TEST_CASE("padding_funcname", "[pattern_formatter]") {
    spdlog::sinks::test_sink_st test_sink;

    const char *pattern = "%v [%10!]";
    auto formatter = std::unique_ptr<spdlog::formatter>(new spdlog::pattern_formatter(pattern));
    test_sink.set_formatter(std::move(formatter));

    spdlog::details::log_msg msg1{spdlog::source_loc{"ignored", 1, "func"}, "test_logger",
                                  spdlog::level::info, "message"};
    test_sink.log(msg1);
    REQUIRE(test_sink.lines()[0] == "message [      func]");

    spdlog::details::log_msg msg2{spdlog::source_loc{"ignored", 1, "func567890123"}, "test_logger",
                                  spdlog::level::info, "message"};
    test_sink.log(msg2);
    REQUIRE(test_sink.lines()[1] == "message [func567890123]");
}

TEST_CASE("clone-default-formatter", "[pattern_formatter]") {
    auto formatter_1 = std::make_shared<spdlog::pattern_formatter>();
    auto formatter_2 = formatter_1->clone();
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "some message");

    memory_buf_t formatted_1;
    memory_buf_t formatted_2;
    formatter_1->format(msg, formatted_1);
    formatter_2->format(msg, formatted_2);

    REQUIRE(to_string_view(formatted_1) == to_string_view(formatted_2));
}

TEST_CASE("clone-default-formatter2", "[pattern_formatter]") {
    auto formatter_1 = std::make_shared<spdlog::pattern_formatter>("%+");
    auto formatter_2 = formatter_1->clone();
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "some message");

    memory_buf_t formatted_1;
    memory_buf_t formatted_2;
    formatter_1->format(msg, formatted_1);
    formatter_2->format(msg, formatted_2);

    REQUIRE(to_string_view(formatted_1) == to_string_view(formatted_2));
}

TEST_CASE("clone-formatter", "[pattern_formatter]") {
    auto formatter_1 = std::make_shared<spdlog::pattern_formatter>("%D %X [%] [%n] %v");
    auto formatter_2 = formatter_1->clone();
    std::string logger_name = "test";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "some message");

    memory_buf_t formatted_1;
    memory_buf_t formatted_2;
    formatter_1->format(msg, formatted_1);
    formatter_2->format(msg, formatted_2);

    REQUIRE(to_string_view(formatted_1) == to_string_view(formatted_2));
}

TEST_CASE("clone-formatter-2", "[pattern_formatter]") {
    using spdlog::pattern_time_type;
    auto formatter_1 = std::make_shared<spdlog::pattern_formatter>(
        "%D %X [%] [%n] %v", pattern_time_type::utc, "xxxxxx\n");
    auto formatter_2 = formatter_1->clone();
    std::string logger_name = "test2";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "some message");

    memory_buf_t formatted_1;
    memory_buf_t formatted_2;
    formatter_1->format(msg, formatted_1);
    formatter_2->format(msg, formatted_2);

    REQUIRE(to_string_view(formatted_1) == to_string_view(formatted_2));
}

class custom_test_flag : public spdlog::custom_flag_formatter {
public:
    explicit custom_test_flag(std::string txt)
        : some_txt{std::move(txt)} {}

    void format(const spdlog::details::log_msg &,
                const std::tm &tm,
                spdlog::memory_buf_t &dest) override {
        if (some_txt == "throw_me") {
            throw spdlog::spdlog_ex("custom_flag_exception_test");
        } else if (some_txt == "time") {
            auto formatted = spdlog::fmt_lib::format("{:d}:{:02d}{:s}", tm.tm_hour % 12, tm.tm_min,
                                                     tm.tm_hour / 12 ? "PM" : "AM");
            dest.append(formatted.data(), formatted.data() + formatted.size());
            return;
        }
        some_txt = std::string(padinfo_.width_, ' ') + some_txt;
        dest.append(some_txt.data(), some_txt.data() + some_txt.size());
    }
    spdlog::details::padding_info get_padding_info() { return padinfo_; }

    std::string some_txt;

    std::unique_ptr<custom_flag_formatter> clone() const override {
        return spdlog::details::make_unique<custom_test_flag>(some_txt);
    }
};
// test clone with custom flag formatters
TEST_CASE("clone-custom_formatter", "[pattern_formatter]") {
    auto formatter_1 = std::make_shared<spdlog::pattern_formatter>();
    formatter_1->add_flag<custom_test_flag>('t', "custom_output").set_pattern("[%n] [%t] %v");
    auto formatter_2 = formatter_1->clone();
    std::string logger_name = "logger-name";
    spdlog::details::log_msg msg(logger_name, spdlog::level::info, "some message");

    memory_buf_t formatted_1;
    memory_buf_t formatted_2;
    formatter_1->format(msg, formatted_1);
    formatter_2->format(msg, formatted_2);

    auto expected = spdlog::fmt_lib::format("[logger-name] [custom_output] some message{}",
                                            spdlog::details::os::default_eol);

    REQUIRE(to_string_view(formatted_1) == expected);
    REQUIRE(to_string_view(formatted_2) == expected);
}

//
// Test source location formatting
//

#ifdef _WIN32
static const char *const test_path = "\\a\\b\\c/myfile.cpp";
#else
static const char *const test_path = "/a/b//myfile.cpp";
#endif

TEST_CASE("short filename formatter-1", "[pattern_formatter]") {
    spdlog::pattern_formatter formatter("%s", spdlog::pattern_time_type::local, "");
    memory_buf_t formatted;
    std::string logger_name = "logger-name";
    spdlog::source_loc source_loc{test_path, 123, "some_func()"};
    spdlog::details::log_msg msg(source_loc, "logger-name", spdlog::level::info, "Hello");
    formatter.format(msg, formatted);

    REQUIRE(to_string_view(formatted) == "myfile.cpp");
}

TEST_CASE("short filename formatter-2", "[pattern_formatter]") {
    spdlog::pattern_formatter formatter("%s:%#", spdlog::pattern_time_type::local, "");
    memory_buf_t formatted;
    std::string logger_name = "logger-name";
    spdlog::source_loc source_loc{"myfile.cpp", 123, "some_func()"};
    spdlog::details::log_msg msg(source_loc, "logger-name", spdlog::level::info, "Hello");
    formatter.format(msg, formatted);

    REQUIRE(to_string_view(formatted) == "myfile.cpp:123");
}

TEST_CASE("short filename formatter-3", "[pattern_formatter]") {
    spdlog::pattern_formatter formatter("%s %v", spdlog::pattern_time_type::local, "");
    memory_buf_t formatted;
    std::string logger_name = "logger-name";
    spdlog::source_loc source_loc{"", 123, "some_func()"};
    spdlog::details::log_msg msg(source_loc, "logger-name", spdlog::level::info, "Hello");
    formatter.format(msg, formatted);

    REQUIRE(to_string_view(formatted) == " Hello");
}

TEST_CASE("full filename formatter", "[pattern_formatter]") {
    spdlog::pattern_formatter formatter("%g", spdlog::pattern_time_type::local, "");
    memory_buf_t formatted;
    std::string logger_name = "logger-name";
    spdlog::source_loc source_loc{test_path, 123, "some_func()"};
    spdlog::details::log_msg msg(source_loc, "logger-name", spdlog::level::info, "Hello");
    formatter.format(msg, formatted);

    REQUIRE(to_string_view(formatted) == test_path);
}

TEST_CASE("custom flags", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->add_flag<custom_test_flag>('t', "custom1")
        .add_flag<custom_test_flag>('u', "custom2")
        .set_pattern("[%n] [%t] [%u] %v");

    memory_buf_t formatted;

    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted);
    auto expected = spdlog::fmt_lib::format("[logger-name] [custom1] [custom2] some message{}",
                                            spdlog::details::os::default_eol);

    REQUIRE(to_string_view(formatted) == expected);
}

TEST_CASE("custom flags-padding", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->add_flag<custom_test_flag>('t', "custom1")
        .add_flag<custom_test_flag>('u', "custom2")
        .set_pattern("[%n] [%t] [%5u] %v");

    memory_buf_t formatted;

    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted);
    auto expected = spdlog::fmt_lib::format("[logger-name] [custom1] [     custom2] some message{}",
                                            spdlog::details::os::default_eol);

    REQUIRE(to_string_view(formatted) == expected);
}

TEST_CASE("custom flags-exception", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->add_flag<custom_test_flag>('t', "throw_me")
        .add_flag<custom_test_flag>('u', "custom2")
        .set_pattern("[%n] [%t] [%u] %v");

    memory_buf_t formatted;
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    CHECK_THROWS_AS(formatter->format(msg, formatted), spdlog::spdlog_ex);
}

TEST_CASE("override need_localtime", "[pattern_formatter]") {
    auto formatter =
        std::make_shared<spdlog::pattern_formatter>(spdlog::pattern_time_type::local, "\n");
    formatter->add_flag<custom_test_flag>('t', "time").set_pattern("%t> %v");

    {
        memory_buf_t formatted;
        spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                     "some message");
        formatter->format(msg, formatted);
        REQUIRE(to_string_view(formatted) == "0:00AM> some message\n");
    }

    {
        formatter->need_localtime();

        auto now_tm = spdlog::details::os::localtime();
        std::stringstream oss;
        oss << (now_tm.tm_hour % 12) << ":" << std::setfill('0') << std::setw(2) << now_tm.tm_min
            << (now_tm.tm_hour / 12 ? "PM" : "AM") << "> some message\n";

        memory_buf_t formatted;
        spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                     "some message");
        formatter->format(msg, formatted);
        REQUIRE(to_string_view(formatted) == oss.str());
    }
}

#ifndef SPDLOG_NO_TLS
TEST_CASE("mdc formatter test-1", "[pattern_formatter]") {
    spdlog::mdc::put("mdc_key_1", "mdc_value_1");
    spdlog::mdc::put("mdc_key_2", "mdc_value_2");

    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->set_pattern("[%n] [%l] [%&] %v");

    memory_buf_t formatted;
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted);

    auto expected = spdlog::fmt_lib::format(
        "[logger-name] [info] [mdc_key_1:mdc_value_1 mdc_key_2:mdc_value_2] some message{}",
        spdlog::details::os::default_eol);
    REQUIRE(to_string_view(formatted) == expected);

    SECTION("Tear down") { spdlog::mdc::clear(); }
}

TEST_CASE("mdc formatter value update", "[pattern_formatter]") {
    spdlog::mdc::put("mdc_key_1", "mdc_value_1");
    spdlog::mdc::put("mdc_key_2", "mdc_value_2");

    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->set_pattern("[%n] [%l] [%&] %v");

    memory_buf_t formatted_1;
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted_1);

    auto expected = spdlog::fmt_lib::format(
        "[logger-name] [info] [mdc_key_1:mdc_value_1 mdc_key_2:mdc_value_2] some message{}",
        spdlog::details::os::default_eol);

    REQUIRE(to_string_view(formatted_1) == expected);

    spdlog::mdc::put("mdc_key_1", "new_mdc_value_1");
    memory_buf_t formatted_2;
    formatter->format(msg, formatted_2);
    expected = spdlog::fmt_lib::format(
        "[logger-name] [info] [mdc_key_1:new_mdc_value_1 mdc_key_2:mdc_value_2] some message{}",
        spdlog::details::os::default_eol);

    REQUIRE(to_string_view(formatted_2) == expected);

    SECTION("Tear down") { spdlog::mdc::clear(); }
}

TEST_CASE("mdc different threads", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->set_pattern("[%n] [%l] [%&] %v");
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");

    memory_buf_t formatted_2;

    auto lambda_1 = [formatter, msg]() {
        spdlog::mdc::put("mdc_key", "thread_1_id");
        memory_buf_t formatted;
        formatter->format(msg, formatted);

        auto expected =
            spdlog::fmt_lib::format("[logger-name] [info] [mdc_key:thread_1_id] some message{}",
                                    spdlog::details::os::default_eol);

        REQUIRE(to_string_view(formatted) == expected);
    };

    auto lambda_2 = [formatter, msg]() {
        spdlog::mdc::put("mdc_key", "thread_2_id");
        memory_buf_t formatted;
        formatter->format(msg, formatted);

        auto expected =
            spdlog::fmt_lib::format("[logger-name] [info] [mdc_key:thread_2_id] some message{}",
                                    spdlog::details::os::default_eol);

        REQUIRE(to_string_view(formatted) == expected);
    };

    std::thread thread_1(lambda_1);
    std::thread thread_2(lambda_2);

    thread_1.join();
    thread_2.join();

    SECTION("Tear down") { spdlog::mdc::clear(); }
}

TEST_CASE("mdc remove key", "[pattern_formatter]") {
    spdlog::mdc::put("mdc_key_1", "mdc_value_1");
    spdlog::mdc::put("mdc_key_2", "mdc_value_2");
    spdlog::mdc::remove("mdc_key_1");

    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->set_pattern("[%n] [%l] [%&] %v");

    memory_buf_t formatted;
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted);

    auto expected =
        spdlog::fmt_lib::format("[logger-name] [info] [mdc_key_2:mdc_value_2] some message{}",
                                spdlog::details::os::default_eol);
    REQUIRE(to_string_view(formatted) == expected);

    SECTION("Tear down") { spdlog::mdc::clear(); }
}

TEST_CASE("mdc empty", "[pattern_formatter]") {
    auto formatter = std::make_shared<spdlog::pattern_formatter>();
    formatter->set_pattern("[%n] [%l] [%&] %v");

    memory_buf_t formatted;
    spdlog::details::log_msg msg(spdlog::source_loc{}, "logger-name", spdlog::level::info,
                                 "some message");
    formatter->format(msg, formatted);

    auto expected = spdlog::fmt_lib::format("[logger-name] [info] [] some message{}",
                                            spdlog::details::os::default_eol);
    REQUIRE(to_string_view(formatted) == expected);

    SECTION("Tear down") { spdlog::mdc::clear(); }
}
#endif
