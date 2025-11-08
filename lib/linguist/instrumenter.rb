module Linguist
  class BasicInstrumenter
    attr_reader :detected_info

    def initialize
      @detected_info = {}
    end

    def instrument(name, payload = {})
      if name == "linguist.detected" && payload[:blob]
        strategies = payload[:strategies] || [payload[:strategy]].compact
        strategy_names = strategies.map { |s| s.name.split("::").last }.join("+")
        blob_name = payload[:blob].name.to_s.force_encoding("UTF-8").scrub
        @detected_info[blob_name] = {
          strategy: strategy_names,
          language: payload[:language]&.name
        }
      end
      yield if block_given?
    end
  end
end
