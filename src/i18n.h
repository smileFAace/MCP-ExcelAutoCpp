#ifndef I18N_H
#define I18N_H

#include <string>
#include <vector>
#include <unordered_map>
#include <sstream> // Added for std::stringstream
#include "json.hpp" // Corrected include path for nlohmann/json

// Use the correct namespace for the JSON library
namespace nlohmann { using json = basic_json<>; }

namespace i18n {

class I18nManager {
public:
    // Loads language data from a JSON file.
    // Returns true on success, false otherwise.
    bool loadLanguage(const std::string& langCode, const std::string& filePath);

    // Loads language data from a JSON string.
    // Returns true on success, false otherwise.
    bool loadLanguageFromString(const std::string& langCode, const std::string& jsonContent);

    // Sets the current language to use for translations.
    // Returns true if the language was loaded previously, false otherwise.
    bool setLanguage(const std::string& langCode);

    // Gets the translation for a given key in the current language.
    // Returns the key itself if the translation is not found.
    std::string get(const std::string& key) const;

    // Gets the translation for a given key and formats it with arguments.
    // Uses simple placeholder replacement like {0}, {1}, etc.
    template<typename... Args>
    std::string get(const std::string& key, Args&&... args) const {
        std::string formatString = get(key);
        return format(formatString, std::forward<Args>(args)...);
    }

    // Returns the currently set language code.
    const std::string& getCurrentLanguage() const;

    // Gets the singleton instance of the I18nManager.
    static I18nManager& getInstance();

private:
    I18nManager() = default; // Private constructor for singleton
    ~I18nManager() = default;
    I18nManager(const I18nManager&) = delete;
    I18nManager& operator=(const I18nManager&) = delete;

    // Helper function for string formatting.
    template<typename T, typename... Args>
    std::string format(std::string formatString, T&& value, Args&&... args) const {
        size_t pos = formatString.find("{" + std::to_string(placeholderIndex++) + "}");
        if (pos != std::string::npos) {
            std::stringstream ss;
            ss << value;
            formatString.replace(pos, 3 + std::to_string(placeholderIndex - 1).length(), ss.str());
            return format(formatString, std::forward<Args>(args)...);
        }
        // Reset index if we still have arguments but no more placeholders
        placeholderIndex = 0;
        return formatString; // Return partially formatted string if placeholders run out
    }

    // Base case for the recursive format function.
    std::string format(std::string formatString) const {
        placeholderIndex = 0; // Reset index for the next call
        return formatString;
    }


    std::unordered_map<std::string, nlohmann::json> languages_; // Stores loaded language data (langCode -> json object)
    std::string currentLangCode_;                         // Currently active language code
    mutable size_t placeholderIndex = 0;                  // Mutable index for formatting placeholders
};

// Global convenience function to get translations.
inline std::string t(const std::string& key) {
    return I18nManager::getInstance().get(key);
}

// Global convenience function to get formatted translations.
template<typename... Args>
inline std::string t(const std::string& key, Args&&... args) {
    return I18nManager::getInstance().get(key, std::forward<Args>(args)...);
}

} // namespace i18n

#endif // I18N_H