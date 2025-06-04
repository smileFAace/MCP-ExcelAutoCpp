#include "i18n.h"
#include <fstream>
#include <sstream>
#include <iostream> // For std::cerr in case spdlog is not yet initialized
#include "spdlog/spdlog.h" // For logging errors

namespace i18n {

// Initialize the singleton instance
I18nManager& I18nManager::getInstance() {
    static I18nManager instance;
    return instance;
}

// Load language data from a JSON file
bool I18nManager::loadLanguage(const std::string& langCode, const std::string& filePath) {
    std::ifstream fileStream(filePath);
    if (!fileStream.is_open()) {
        spdlog::error("i18n: Failed to open language file: {}", filePath);
        // Fallback to cerr if spdlog might not be ready
        // std::cerr << "i18n: Failed to open language file: " << filePath << std::endl;
        return false;
    }

    try {
        nlohmann::json langData = nlohmann::json::parse(fileStream);
        languages_[langCode] = langData;
        spdlog::info("i18n: Successfully loaded language file '{}' for language code '{}'", filePath, langCode);
        // Set the first loaded language as the default/current one if none is set
        if (currentLangCode_.empty()) {
            setLanguage(langCode);
        }
        return true;
    } catch (const nlohmann::json::parse_error& e) {
        spdlog::error("i18n: Failed to parse language file '{}': {}", filePath, e.what());
        // std::cerr << "i18n: Failed to parse language file '" << filePath << "': " << e.what() << std::endl;
        return false;
    } catch (const std::exception& e) {
        spdlog::error("i18n: An unexpected error occurred while loading language file '{}': {}", filePath, e.what());
        // std::cerr << "i18n: An unexpected error occurred while loading language file '" << filePath << "': " << e.what() << std::endl;
        return false;
    }
}

// Load language data from a JSON string
bool I18nManager::loadLanguageFromString(const std::string& langCode, const std::string& jsonContent) {
    try {
        nlohmann::json langData = nlohmann::json::parse(jsonContent);
        languages_[langCode] = langData;
        spdlog::info("i18n: Successfully loaded language string for language code '{}'", langCode);
        // Set the first loaded language as the default/current one if none is set
        if (currentLangCode_.empty()) {
            setLanguage(langCode);
        }
        return true;
    } catch (const nlohmann::json::parse_error& e) {
        spdlog::error("i18n: Failed to parse language string for code '{}': {}", langCode, e.what());
        return false;
    } catch (const std::exception& e) {
        spdlog::error("i18n: An unexpected error occurred while loading language string for code '{}': {}", langCode, e.what());
        return false;
    }
}


// Set the current language
bool I18nManager::setLanguage(const std::string& langCode) {
    if (languages_.count(langCode)) {
        currentLangCode_ = langCode;
        spdlog::info("i18n: Set current language to '{}'", langCode);
        return true;
    } else {
        spdlog::warn("i18n: Attempted to set language to '{}', but it was not loaded.", langCode);
        return false;
    }
}

// Get the translation for a key
std::string I18nManager::get(const std::string& key) const {
    if (currentLangCode_.empty() || languages_.find(currentLangCode_) == languages_.end()) {
        spdlog::warn("i18n: No language set or current language '{}' not loaded. Returning key '{}'.", currentLangCode_, key);
        return key; // Return the key itself if no language is set or loaded
    }

    const auto& langData = languages_.at(currentLangCode_);
    try {
        // Navigate nested keys if necessary (e.g., "errors.file_not_found")
        nlohmann::json current = langData;
        std::string segment;
        std::stringstream ss(key);
        bool first = true;
        while (std::getline(ss, segment, '.')) {
            if (!current.contains(segment)) {
                 spdlog::warn("i18n: Translation key '{}' not found in language '{}'.", key, currentLangCode_);
                 return key; // Return the original key if not found
            }
            current = current[segment];
            first = false;
        }


        if (current.is_string()) {
            return current.get<std::string>();
        } else {
             spdlog::warn("i18n: Translation key '{}' in language '{}' is not a string.", key, currentLangCode_);
            return key; // Return key if the found value isn't a string
        }
    } catch (const nlohmann::json::exception& e) {
        spdlog::warn("i18n: Error accessing translation key '{}' in language '{}': {}", key, currentLangCode_, e.what());
        return key; // Return the key in case of JSON access errors
    } catch (const std::exception& e) {
         spdlog::error("i18n: Unexpected error getting translation for key '{}' in language '{}': {}", key, currentLangCode_, e.what());
         return key;
    }
}

// Get the current language code
const std::string& I18nManager::getCurrentLanguage() const {
    return currentLangCode_;
}

// Note: The template format functions are defined inline in the header (i18n.h)

} // namespace i18n